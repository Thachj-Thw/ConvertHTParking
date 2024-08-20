def convert8to10(str_number: str) -> str:
    if len(str_number) > 8:
        str_number = str_number[:8]
    number = int(str_number)
    first = number // 100_000
    last = number % 100_000
    return str(first * 65536 + last).rjust(10, "0")
