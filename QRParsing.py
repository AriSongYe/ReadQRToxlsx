def parsingQR(self, input):
    parsed_values = input.split('-')

    if len(parsed_values) == 2:
        first_number = int(parsed_values[0])
    second_number = int(parsed_values[1])