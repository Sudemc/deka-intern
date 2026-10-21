"""Benchmark test suite for Batch 3: Visual Output Devices (DSP - 24 Displays)."""

from pathlib import Path
import pytest
from bom_enricher.comparison import load_rulebook
from bom_enricher.datasheet_analyzer import analyze_datasheet


@pytest.fixture(scope="module")
def rulebook():
    return load_rulebook(Path("data/rules.xlsx"))


def test_batch3_dsp_displays(rulebook):
    # 1. Microtech MTB0802A2
    res = analyze_datasheet(b"", part_number="MTB0802A2", rulebook=rulebook, serial_counter="0029")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTD4AA0029-MH"

    # 2. Winstar Display WH0802A1-TFH-ET#
    res = analyze_datasheet(b"", part_number="WH0802A1-TFH-ET#", rulebook=rulebook, serial_counter="0031")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTD4AA0031-WS"

    # 3. YEEBO GR-V0801A-A0
    res = analyze_datasheet(b"", part_number="GR-V0801A-A0", rulebook=rulebook, serial_counter="0057")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE0FE0057-YBO"

    # 4. Liteon LTS-2801AJE
    res = analyze_datasheet(b"", part_number="LTS-2801AJE", rulebook=rulebook, serial_counter="0002")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BA0002-LO"

    # 5. Strong Base E1-1039-AUR-1
    res = analyze_datasheet(b"", part_number="E1-1039-AUR-1", rulebook=rulebook, serial_counter="0007")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BA0007-SB"

    # 6. Foryard FYS-3911BUHR-21
    res = analyze_datasheet(b"", part_number="FYS-3911BUHR-21", rulebook=rulebook, serial_counter="0013")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BA0013-FO"

    # 7. Bright Led BD-A54(6RI / UBRI)
    res = analyze_datasheet(b"", part_number="BD-A54(6RI / UBRI)", rulebook=rulebook, serial_counter="0001")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0001-BR"

    # 8. Avago HDSP-521A
    res = analyze_datasheet(b"", part_number="HDSP-521A", rulebook=rulebook, serial_counter="0005")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0005-AA"

    # 9. Strong Base E2-2039AUR1-P8.3
    res = analyze_datasheet(b"", part_number="E2-2039AUR1-P8.3", rulebook=rulebook, serial_counter="0009")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0009-SB"

    # 10. Wuxi Compul CPS03622BUR-21
    res = analyze_datasheet(b"", part_number="CPS03622BUR-21", rulebook=rulebook, serial_counter="0012")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0012-WX"

    # 11. Strong Base E3-2056ASR2
    res = analyze_datasheet(b"", part_number="E3-2056ASR2", rulebook=rulebook, serial_counter="0014")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0014-SB"

    # 12. Strong Base E3-2056AUR2
    res = analyze_datasheet(b"", part_number="E3-2056AUR2", rulebook=rulebook, serial_counter="0021")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0021-SB"

    # 13. Kinley KL03922WDGBSR
    res = analyze_datasheet(b"", part_number="KL03922WDGBSR", rulebook=rulebook, serial_counter="0023")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0023-KLY"

    # 14. Kinley KL03622WDGBSR
    res = analyze_datasheet(b"", part_number="KL03622WDGBSR", rulebook=rulebook, serial_counter="0027")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BC0027-KLY"

    # 15. Kinley KL03631WDGBSR
    res = analyze_datasheet(b"", part_number="KL03631WDGBSR", rulebook=rulebook, serial_counter="0025")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BE0025-KLY"

    # 16. Assemcorp KL05632WDGAW
    res = analyze_datasheet(b"", part_number="KL05632WDGAW", rulebook=rulebook, serial_counter="0038")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE1BF0038-AS"

    # 17. Semtech 2414A-DW-6
    res = analyze_datasheet(b"", part_number="2414A-DW-6", rulebook=rulebook, serial_counter="0055")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE4BF0055-SM"

    # 18. Semtech 2614A-DW-6
    res = analyze_datasheet(b"", part_number="2614A-DW-6", rulebook=rulebook, serial_counter="0056")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE4BF0056-SM"

    # 19. Foryard FYQ-5643CUW-21
    res = analyze_datasheet(b"", part_number="FYQ-5643CUW-21", rulebook=rulebook, serial_counter="0028")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE4BH0028-FO"

    # 20. Kinley KL05643WDGAW1-0603
    res = analyze_datasheet(b"", part_number="KL05643WDGAW1-0603", rulebook=rulebook, serial_counter="0030")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE4BH0030-KLY"

    # 21. Semtech ALS3819-21W
    res = analyze_datasheet(b"", part_number="ALS3819-21W", rulebook=rulebook, serial_counter="0046")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTE4BH0046-SM"

    # 22. Display Visions EA DOGM162E-A
    res = analyze_datasheet(b"", part_number="EA DOGM162E-A", rulebook=rulebook, serial_counter="0004")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTH3AB0004-DV"

    # 23. Display Visions EA LED55X31-G
    res = analyze_datasheet(b"", part_number="EA LED55X31-G", rulebook=rulebook, serial_counter="0006")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTI5000006-DV"

    # 24. Nextion NX4827P043-011C-Y
    res = analyze_datasheet(b"", part_number="NX4827P043-011C-Y", rulebook=rulebook, serial_counter="0047")
    assert res["material_type"] == "DSP"
    assert res["erp_code"] == "DSPTJ0FC0047-NT"
