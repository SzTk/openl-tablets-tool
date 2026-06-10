from templates import derive_service_name, generate_rules_deploy_xml, generate_rules_xml


def test_derive_service_name_from_pascal_case():
    assert derive_service_name("ShopPolicy.xlsx") == "shop-policy"


def test_derive_service_name_from_multi_word_pascal_case():
    assert derive_service_name("AutoPolicyCalculation.xlsx") == "auto-policy-calculation"


def test_derive_service_name_with_underscore_and_digit():
    assert derive_service_name("ShopPolicy2_draft.xlsx") == "shop-policy2-draft"


def test_generate_rules_xml_contains_service_name_and_filename():
    xml = generate_rules_xml("shop-policy", "ShopPolicy.xlsx")

    assert "<name>shop-policy</name>" in xml
    assert '<rules-root path="ShopPolicy.xlsx"/>' in xml


def test_generate_rules_deploy_xml_contains_service_name():
    xml = generate_rules_deploy_xml("shop-policy")

    assert "<serviceName>shop-policy</serviceName>" in xml
    assert "<url>shop-policy</url>" in xml
    assert "<publisher>RESTFUL</publisher>" in xml
