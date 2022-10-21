import openpyxl
from openpyxl.styles import PatternFill


'''
próximos a llamar (azul): FF0070C0,
fill; on

pendientes (rojo): FFFF0000, FF7030A0

pacientes a llamar (verde): FF00B050, FF7030A0
fill; FF92D050

rellamados (rosa): FFFF00FF
fill; FFFFC000, FFFFFF00

otro: Values must be of type <class 'str'>, None
fill; 00000000
'''

book = openpyxl.load_workbook('./book.xlsx')
sheet = book.active


def all_data(row: int):
    vacuna = sheet[f'E{row}']
    vac_fecha = sheet[f'F{row}']
    desp_fecha = sheet[f'G{row}']
    desparasitante = sheet[f'H{row}']
    print(row, 'FECHA', 'COSO', sep='\t')
    print('%'*80)
    print('fill vacuna:', get_fillRGB(vac_fecha), get_fillRGB(vacuna), sep='\t')
    print('fill desparasitante:', get_fillRGB(desp_fecha), get_fillRGB(desparasitante), sep='\t')
    print('=='*40)
    print('font vacuna:', get_fontRGB(vac_fecha), get_fontRGB(vacuna), sep='\t')
    print('font desparasitante:', get_fontRGB(desp_fecha), get_fontRGB(desparasitante), sep='\t')
    print('&&'*40, '\n')


def get_fontRGB(cell) -> str:
    try:
        rgb = str(cell.font.color.rgb)
    except AttributeError:
        rgb = None
    return rgb


def get_fillRGB(cell) -> str:
    properties = str(cell.fill).split()
    rgbs = filter(lambda x: 'rgb' in x, properties)
    rgb = list(rgbs)[0]
    hex = rgb[5: -2]
    return hex


def all_RGB() -> None:
    row = 1
    vac_fill = {'vac' :set(), 'fecha': set()}
    desp_fill = {'desp' :set(), 'fecha': set()}

    vac_font = {'vac' :set(), 'fecha': set()}
    desp_font = {'desp' :set(), 'fecha': set()}

    while True:
        row += 1
        owner = sheet[f'A{row}']
        if owner.value == None:
            break
        vacuna = sheet[f'E{row}']
        vac_fecha = sheet[f'F{row}']
        desp_fecha = sheet[f'G{row}']
        desparasitante = sheet[f'H{row}']
    
        vac_fill['vac'].add(get_fillRGB(vacuna))
        vac_fill['fecha'].add(get_fillRGB(vac_fecha))

        vac_font['vac'].add(get_fontRGB(vacuna))
        vac_font['fecha'].add(get_fontRGB(vac_fecha))

        desp_fill['desp'].add(get_fillRGB(desparasitante))
        desp_fill['fecha'].add(get_fillRGB(desp_fecha))

        desp_font['desp'].add(get_fontRGB(desparasitante))
        desp_font['fecha'].add(get_fontRGB(desp_fecha))

    print('vac_fill (E F) : ', vac_fill, sep='\n')
    print('vac_font: ', vac_font, sep='\n')

    print('desp_fill (H G): ', desp_fill, sep='\n')
    print('desp_font: ', desp_font, sep='\n')


def to_list(rows: int):
    '''to list the fill and font on a set of rows'''
    for row in range(2, rows):
        cell = sheet[f'F{row}']
        print(row)
        print('-'*20)
        print('font: ', get_fontRGB(cell))
        print('fill: ', get_fillRGB(cell))
        print('_'*30)


def get_row(RGB: str, column: str = 'F') -> str:
    '''search a property with the RGB and return his row and print the field'''
    row = 1
    while True:
        row += 1
        owner = sheet[f'A{row}']
        if owner.value == None:
            return 'No found'
        cell = sheet[column + str(row)]
        if get_fillRGB(cell) == RGB:
            return f'fill at row {row}'
        if get_fontRGB(cell) == RGB:
            return f'font at row {row}'


def run():
    print(get_row('FF7030A0', 'G'))


if __name__ == '__main__':
    run()