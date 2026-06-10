import re


def derive_service_name(excel_filename: str) -> str:
    """Derive a URL-safe service name from an Excel filename.

    "ShopPolicy.xlsx" -> "shop-policy"
    "AutoPolicyCalculation.xlsx" -> "auto-policy-calculation"
    "ShopPolicy2_draft.xlsx" -> "shop-policy2-draft"
    """
    stem = excel_filename.rsplit(".", 1)[0]
    stem = re.sub(r"(?<!^)(?=[A-Z])", "-", stem)
    stem = re.sub(r"[^a-zA-Z0-9]+", "-", stem)
    return stem.strip("-").lower()


def generate_rules_xml(service_name: str, excel_filename: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<project>
  <name>{service_name}</name>
  <modules>
    <module>
      <name>{service_name}</name>
      <rules-root path="{excel_filename}"/>
    </module>
  </modules>
</project>
"""


def generate_rules_deploy_xml(service_name: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<rules-deploy>
  <isProvideRuntimeContext>false</isProvideRuntimeContext>
  <serviceName>{service_name}</serviceName>
  <url>{service_name}</url>
  <publishers>
    <publisher>RESTFUL</publisher>
  </publishers>
</rules-deploy>
"""
