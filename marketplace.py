import json
import boto3
from decimal import Decimal
from typing import Optional, Tuple


def _first_price_usd_from_price_list(price_list_item: str) -> Optional[float]:
    """
    Parse a single PriceList JSON string from GetProducts and return the first USD on-demand price per unit (hour).
    Returns None if not found.
    """
    obj = json.loads(price_list_item)
    terms = obj.get("terms", {}).get("OnDemand", {})
    if not terms:
        return None
    # Grab first term → first priceDimension → USD
    term = next(iter(terms.values()))
    pdims = term.get("priceDimensions", {})
    if not pdims:
        return None
    pd = next(iter(pdims.values()))
    price_str = pd.get("pricePerUnit", {}).get("USD")
    if price_str is None:
        return None
    try:
        return float(Decimal(price_str))
    except Exception:
        return None


def _pricing_client():
    # Pricing API endpoints exist only in us-east-1 and ap-south-1
    return boto3.client("pricing", region_name="us-east-1")


def get_ec2_ondemand_hourly(
    instance_type: str,
    location: str = "US East (N. Virginia)",
    operating_system: str = "RHEL",
    tenancy: str = "Shared",
    pre_installed_sw: str = "NA",
) -> Optional[float]:
    """
    Returns the EC2 On-Demand hourly price for a given instance type, OS, and location.
    Defaults tailored for RHEL on shared tenancy. Returns None if not found.
    """
    client = _pricing_client()
    # Build filters per Pricing Query API
    filters = [
        {"Type": "TERM_MATCH", "Field": "instanceType", "Value": instance_type},
        {"Type": "TERM_MATCH", "Field": "location", "Value": location},
        {"Type": "TERM_MATCH", "Field": "operatingSystem", "Value": operating_system},
        {"Type": "TERM_MATCH", "Field": "tenancy", "Value": tenancy},
        {"Type": "TERM_MATCH", "Field": "preInstalledSw", "Value": pre_installed_sw},
        {"Type": "TERM_MATCH", "Field": "capacitystatus", "Value": "Used"},
        {"Type": "TERM_MATCH", "Field": "termType", "Value": "OnDemand"},
    ]
    resp = client.get_products(ServiceCode="AmazonEC2", Filters=filters, MaxResults=100)
    for item in resp.get("PriceList", []):
        price = _first_price_usd_from_price_list(item)
        if price is not None:
            return price
    # If paginated, you could loop on NextToken here.
    return None


def get_marketplace_ami_hourly(
    product_title_contains: str,
    instance_type: str,
    location: str = "US East (N. Virginia)",
) -> Optional[float]:
    """
    Returns the AWS Marketplace AMI hourly *software* fee for a given AMI (by title match),
    instance type, and location. Returns None if not found.

    This looks for products with:
      - productFamily=Software
      - usagetype=SoftwareUsage:<instance_type>
      - location=<region display name>
      - and product title contains the provided string (case-insensitive match)
    """
    client = _pricing_client()

    # IMPORTANT: Pricing API filters are exact-matches; we cannot 'contains' on title via Filter.
    # Strategy: fetch candidates for the usage type + location, then test product attributes' title client-side.
    filters = [
        {"Type": "TERM_MATCH", "Field": "productFamily", "Value": "Software"},
        {"Type": "TERM_MATCH", "Field": "usagetype", "Value": f"SoftwareUsage:{instance_type}"},
        {"Type": "TERM_MATCH", "Field": "location", "Value": location},
    ]

    resp = client.get_products(ServiceCode="AmazonEC2", Filters=filters, MaxResults=100)
    title_lc = product_title_contains.lower()

    for item in resp.get("PriceList", []):
        obj = json.loads(item)
        # Titles appear in 'product' → 'attributes' under various keys; common is 'softwareType' + 'title'/'productTitle'.
        attrs = obj.get("product", {}).get("attributes", {})
        cand_title = (
            attrs.get("productTitle")
            or attrs.get("title")
            or attrs.get("softwareName")
            or ""
        )
        if title_lc in str(cand_title).lower():
            price = _first_price_usd_from_price_list(item)
            if price is not None:
                return price

    # If you don’t find by title, return the first price found for the usage type (some listings have concise titles)
    for item in resp.get("PriceList", []):
        price = _first_price_usd_from_price_list(item)
        if price is not None:
            return price

    return None


def get_marketplace_ami_total_hourly(
    product_title_contains: str,
    instance_type: str,
    location: str = "US East (N. Virginia)",
    operating_system: str = "RHEL",
) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    """
    Convenience wrapper: returns (ec2_price, marketplace_fee, total).
    """
    ec2_price = get_ec2_ondemand_hourly(
        instance_type=instance_type,
        location=location,
        operating_system=operating_system,
    )
    mkt_fee = get_marketplace_ami_hourly(
        product_title_contains=product_title_contains,
        instance_type=instance_type,
        location=location,
    )
    total = (ec2_price + mkt_fee) if (ec2_price is not None and mkt_fee is not None) else None
    return ec2_price, mkt_fee, total


if __name__ == "__main__":
    # Examples for your two rows (adjust the location string to your target Region in AWS console wording):
    product = "CIS Hardened Image STIG on Red Hat Enterprise Linux 8"

    for itype in ["c5n.4xlarge", "m6i.8xlarge"]:
        ec2, fee, total = get_marketplace_ami_total_hourly(
            product_title_contains=product,
            instance_type=itype,
            location="US East (N. Virginia)",
            operating_system="RHEL",
        )
        print(f"{product} on {itype} in us-east-1:")
        print(f"  EC2 On-Demand (RHEL):   {ec2!r} USD/hr")
        print(f"  Marketplace software:   {fee!r} USD/hr")
        print(f"  ------------------------------")
        print(f"  TOTAL:                  {total!r} USD/hr\n")