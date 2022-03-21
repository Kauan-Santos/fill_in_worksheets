"""Esse módulo tem como objetivo buscar e realizar as funções do módulo functions.py."""


from functions import *


def main(): 
    """Esta função tem o objetivo enviar os parâmetros para as funções,
    para preencher a planilha 'Example.xlsx',
    de acordo com os parâmetros passados para a cada função abaixo.

    Os parâmetros podem substituidos de acordo com sua necessidade e documento a ser preenchido.
    """
    value_in_cell('2000', 'Old one')
    value_in_cell('2003', 'Old one')
    value_in_cell('2008', 'Old one')

    value_is_equal('Hyper-V 2016', 'Hyper-V')

    equal_value_options('AIX 7.1.0.0 TL5', 'AIX 7.1.0.0 TL4', 'AIX 4 ou 5')

    value_in_worksheet('working days', 'A')

    wb_example.save('Example.xlsx')


if __name__ == '__main__':
    main()