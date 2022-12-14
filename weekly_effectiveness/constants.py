COUNTRIES = {}  # Format {geo name(str): country name(str)}
COURIER_TARIFFS = {}  # Format {geo name (str): {'sent': tariff(dec), 'return': tariff(dec)}}
OPEX_GOODS_COST = {}  # Format {geo name (str): {'opex': tariff(dec), 'goods cost': tariff(dec)}}
RATES = {}  # Currency rates in format {geo name(str): rate to eur(dec)}
DELIVERY_VAT_RATES = {}  # VAT rate for delivery services in format {geo name(str): rate(int, dec)}
REDUCED_VAT_RATE = ()  # List of geos with reduced VAT rates in format (geo name(str), geo name(str), ...)
VAT_RATES_DICT_FROM_BL = {}  # VAT rates for products in format {geo name(str): {product name(str): VAT rate(dec)}}
TOTAL_ROWS = {}  # Gets from the data_passed.json file with eu_total_compiler.py file
DATES = {}  # Gets from the data_passed.json file with eu_total_compiler.py file
