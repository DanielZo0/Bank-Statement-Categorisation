from common_categorization import parse_date_smart, parse_number


def test_parse_number_handles_eu_and_us_formats():
    assert parse_number("1,234.56") == 1234.56
    assert parse_number("1.234,56") == 1234.56
    assert parse_number("(123.45)") == -123.45
    assert parse_number("123-") == -123.0


def test_parse_date_smart_parses_supported_formats():
    assert parse_date_smart("2025-09-30").strftime("%Y-%m-%d") == "2025-09-30"
    assert parse_date_smart("30/09/2025").strftime("%Y-%m-%d") == "2025-09-30"
    assert parse_date_smart("30 September 2025").strftime("%Y-%m-%d") == "2025-09-30"
