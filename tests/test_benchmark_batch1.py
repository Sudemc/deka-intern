import pytest
from bom_enricher.datasheet_analyzer import analyze_datasheet
from bom_enricher.comparison import load_rulebook


@pytest.fixture(scope="module")
def rulebook():
    return load_rulebook("data/rules.xlsx")


def test_batch1_adp_adapters(rulebook):
    # 1. Simsukian SK03T1200200V
    res = analyze_datasheet(b"", part_number="SK03T1200200V", rulebook=rulebook)
    assert res["material_type"] == "ADP"
    assert res["erp_code"] == "ADPAA00A301011-SMK"
    assert "AC/DC Duvar Tipi Adaptör" in res["rich_description"]
    assert "12VDC" in res["rich_description"]
    assert "2A" in res["rich_description"]
    assert "SIMSUKIAN" in res["rich_description"]

    # 2. Yıldırım CYS 1206-M
    res = analyze_datasheet(b"", part_number="CYS 1206-M", serial_counter="0001", rulebook=rulebook)
    assert res["material_type"] == "ADP"
    assert res["erp_code"] == "ADPAA00A100001-YI"
    assert "Yıldırım Adaptör" in res["rich_description"]

    # 3. Kaizen 20163
    res = analyze_datasheet(b"", part_number="20163", serial_counter="000014", rulebook=rulebook)
    assert res["material_type"] == "ADP"
    assert res["erp_code"] == "ADPAA00A000014-KZ"
    assert "Kaizen" in res["rich_description"]


def test_batch1_ant_antennas(rulebook):
    # 1. TE Connectivity 2108857-4
    res = analyze_datasheet(b"", part_number="2108857-4", serial_counter="000020", rulebook=rulebook)
    assert res["material_type"] == "ANT"
    assert res["erp_code"] == "ANTT3000020-TC"
    assert "TE Connectivity" in res["rich_description"]

    # 2. Ezurio MAF95310
    res = analyze_datasheet(b"", part_number="MAF95310", serial_counter="000021", rulebook=rulebook)
    assert res["material_type"] == "ANT"
    assert res["erp_code"] == "ANTT3000021-EZU"
    assert "Ezurio" in res["rich_description"]

    # 3. Ezurio EMF2449A1-10UFL
    res = analyze_datasheet(b"", part_number="EMF2449A1-10UFL", serial_counter="000022", rulebook=rulebook)
    assert res["material_type"] == "ANT"
    assert res["erp_code"] == "ANTT3000022-EZU"
    assert "Ezurio" in res["rich_description"]

    # 4. TE Connectivity 2108857-1
    res = analyze_datasheet(b"", part_number="2108857-1", serial_counter="000023", rulebook=rulebook)
    assert res["material_type"] == "ANT"
    assert res["erp_code"] == "ANTT3000023-TC"
    assert "TE Connectivity" in res["rich_description"]


def test_batch1_bah_battery_holders(rulebook):
    # 1. Kls L-KLS5-CR2032-09
    res = analyze_datasheet(b"", part_number="L-KLS5-CR2032-09", serial_counter="00033", rulebook=rulebook)
    assert res["material_type"] == "BAH"
    assert res["erp_code"] == "BAHS1000033-KL"
    assert "CR2032 Pil Yuvası" in res["rich_description"]
    assert "Kls" in res["rich_description"]

    # 2. Kls L-KLS5-CR2032-11
    res = analyze_datasheet(b"", part_number="L-KLS5-CR2032-11", serial_counter="00035", rulebook=rulebook)
    assert res["material_type"] == "BAH"
    assert res["erp_code"] == "BAHS1000035-KL"

    # 3. Connfly DS1092-11-N8S
    res = analyze_datasheet(b"", part_number="DS1092-11-N8S", serial_counter="00036", rulebook=rulebook)
    assert res["material_type"] == "BAH"
    assert res["erp_code"] == "BAHS1000036-CF"
    assert "Connfly" in res["rich_description"]

    # 4. Kls L-KLS5-SP-073-R
    res = analyze_datasheet(b"", part_number="L-KLS5-SP-073-R", serial_counter="00076", rulebook=rulebook)
    assert res["material_type"] == "BAH"
    assert res["erp_code"] == "BAHS1100076-KL"
    assert "Kls" in res["rich_description"]

    # 5. Keystone Electronics 55
    res = analyze_datasheet(b"", part_number="55", material_type_override="BAH", serial_counter="00056", rulebook=rulebook)
    assert res["material_type"] == "BAH"
    assert res["erp_code"] == "BAHS1300056-KYS"
    assert "Keystone" in res["rich_description"]


def test_batch1_cap_capacitors(rulebook):
    # 1. Murata Automotive GRT155R61C105KE01D (1uF 16V 10% X5R 0402)
    res = analyze_datasheet(b"", part_number="GRT155R61C105KE01D", rulebook=rulebook)
    assert res["material_type"] == "CAP"
    assert res["erp_code"] == "CAPSS11004B1KI000OB-MT"
    assert "0402" in res["rich_description"]
    assert "1uF" in res["rich_description"]
    assert "16V" in res["rich_description"]
    assert "X5R" in res["rich_description"]
    assert "Automotive" in res["rich_description"]
    assert "Murata" in res["rich_description"]

    # 2. Walsin 0402B102J160CT (1000pF 16V 5% X7R 0402)
    res = analyze_datasheet(b"", part_number="0402B102J160CT", rulebook=rulebook)
    assert res["material_type"] == "CAP"
    assert res["erp_code"] == "CAPSS11008B1JA000GA-WLS"
    assert "0402" in res["rich_description"]
    assert "1000pF" in res["rich_description"]
    assert "16V" in res["rich_description"]
    assert "X7R" in res["rich_description"]
    assert "Walsin" in res["rich_description"]

    # 3. Yageo CC0402JRX7R7BB102 (1000pF 16V 5% X7R 0402)
    res = analyze_datasheet(b"", part_number="CC0402JRX7R7BB102", rulebook=rulebook)
    assert res["material_type"] == "CAP"
    assert res["erp_code"] == "CAPSS11008B1JA000GA-YG"
    assert "Yageo" in res["rich_description"]

    # 4. Samsung CL05B104KO5VPND (0.1uF 16V 10% X7R Automotive 0402)
    res = analyze_datasheet(b"", part_number="CL05B104KO5VPND", rulebook=rulebook)
    assert res["material_type"] == "CAP"
    assert res["erp_code"] == "CAPSS11009B1KA000GB-SS"
    assert "Samsung" in res["rich_description"]

    # 5. Murata GRM155R71C103KA01D (10000pF 16V 10% X7R 0402)
    res = analyze_datasheet(b"", part_number="GRM155R71C103KA01D", rulebook=rulebook)
    assert res["material_type"] == "CAP"
    assert res["erp_code"] == "CAPSS11010B1KA000GA-MT"
    assert "Murata" in res["rich_description"]
