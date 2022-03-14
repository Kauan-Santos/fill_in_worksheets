"""Esse módulo tem como objetivo organizar as funções de preenchimento de planilhas."""

from openpyxl import load_workbook

wb_example = load_workbook('Example')
wb_conditions = load_workbook('Conditions')

ws1_example = wb_example['Plan1']
ws1_conditions = wb_conditions['Plan1']


def value_in_cell(search, status):
    """Está função tem o objetivo de retornar se o valor é igual ou diferente de 5

        parâmetros = 1 (Númerico do tipo inteiro)
    """
    for cell in ws1_example['A']:
        if search in str(cell.value) and ws1_example['C' + str(cell.row) == None]:
            ws1_example['C' + str(cell.row)] = status
        if search in str(cell.value) and ws1_example['C' + str(cell.row) != None]:
            print({search} + 'Value status populated with: ' + ws1_example['C' + str(cell.row)])


def value_is_equal(search, status):


    for cell in ws1_example['B']:    
        if cell.value == search and ws1_example['C' + str(cell.row)] == None:
            ws1_example['C' + str(cell.row)] = status


def equal_value_options(value1, value2, status):
    for cell in ws1_example['B']:
        if cell.value == value1 or cell.value == value2:
            ws1_example['C' + str(cell.row)] = status


def value_in_worksheet(status):
    #define
    pass
