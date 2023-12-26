import requests
from os import path
from colorama import Fore, init
from progress.bar import Bar
from time import sleep
import xlsxwriter


class WalletCheckError(Exception):
    pass


def check_wallet(address):
    try:
        params = {'address': address}
        r = requests.get('https://starkrocket.xyz/api/check_wallet', params=params)

        if r.status_code != 200:
            return WalletCheckError(f'Не удалось проверить кошелёк: status_code = {r.status_code}')

        try:
            r = r.json()
        except requests.exceptions.JSONDecodeError:
            return WalletCheckError(f'Не удалось проверить кошелёк: сервер вернул неверный ответ')

        return r['result']
    except Exception as e:
        return WalletCheckError(str(e))


def check_wallets(wallets_list):
    bar = Bar(message='Проверка', max=len(wallets_list))
    bar.start()

    result = []

    for wallet in wallets_list:
        r = check_wallet(wallet)
        if isinstance(r, WalletCheckError):
            data = {'status': False, 'address': wallet, 'data': str(r)}
        else:
            data = {'status': True, 'data': r}
        result.append(data)
        bar.next()

        sleep(0.5)

    bar.finish()

    return result


def create_report(data, report_type):
    if report_type == 'txt':
        if path.exists('report.txt'):
            for i in range(1, 999):
                if path.exists(f'report_{i}.txt'):
                    continue
                name = f'report_{i}.txt'
                break
        else:
            name = f'report.txt'

        text = ''

        for wallet in data:
            row = ''

            if wallet['status']:
                row += wallet['data']['address'] + '\n'
                row += f'\tПериоды транзакций: {", ".join([str(i) for i in wallet["data"]["criteria"]["transactions_over_time"]]) if wallet["data"]["criteria"]["transactions_over_time"] else "НЕТ"}\n'
                row += f'\tКоличество транзакций: {", ".join([str(i) for i in wallet["data"]["criteria"]["transactions_frequency"]]) if wallet["data"]["criteria"]["transactions_frequency"] else "НЕТ"}\n'
                row += f'\tКоличество контрактов: {", ".join([str(i) for i in wallet["data"]["criteria"]["contracts_variety"]]) if wallet["data"]["criteria"]["contracts_variety"] else "НЕТ"}\n'
                row += f'\tОбъём транзакций: {", ".join([str(i) for i in wallet["data"]["criteria"]["transaction_volume"]]) if wallet["data"]["criteria"]["transaction_volume"] else "НЕТ"}\n'
                row += f'\tОбъём бриджа: {", ".join([str(i) for i in wallet["data"]["criteria"]["bridge_volume"]]) if wallet["data"]["criteria"]["bridge_volume"] else "НЕТ"}\n'
                row += f'\tПоинты: {wallet["data"]["points"]}\n'
                row += f'\tДоступность дропа: {"ДА" if wallet["data"]["eligible"] else "НЕТ"}\n\n'
            else:
                row += wallet['address'] + '\n'
                row += f'\tОшибка проверки: {wallet["data"]}\n\n'

            text += row

        with open(name, 'w', encoding='utf-8') as file:
            file.write(text)
            file.close()

    elif report_type == 'xlsx':
        if path.exists('report.xlsx'):
            for i in range(1, 999):
                if path.exists(f'report_{i}.xlsx'):
                    continue
                name = f'report_{i}.xlsx'
                break
        else:
            name = f'report.xlsx'

        workbook = xlsxwriter.Workbook(name)
        worksheet = workbook.add_worksheet('Результат')

        table_head = workbook.add_format()
        table_head.bold = True
        table_head.bg_color = '#F8CBAD'
        table_head.set_text_h_align(2)
        table_head.set_text_v_align(2)
        table_head.set_border()

        border_format = workbook.add_format()
        border_format.set_border()
        border_format.set_text_h_align(2)
        border_format.set_text_v_align(2)

        wrap_border_format = workbook.add_format()
        wrap_border_format.set_border()
        wrap_border_format.set_text_h_align(2)
        wrap_border_format.set_text_v_align(2)
        wrap_border_format.text_wrap = True

        success_format = workbook.add_format()
        success_format.set_border()
        success_format.bg_color = '#92D050'
        success_format.set_text_h_align(2)
        success_format.set_text_v_align(2)

        unsuccess_format = workbook.add_format()
        unsuccess_format.set_border()
        unsuccess_format.bg_color = '#FF5B5B'
        unsuccess_format.set_text_h_align(2)
        unsuccess_format.set_text_v_align(2)

        headers = ['Адрес', 'Транзакции за периоды', 'Количество транзакций', 'Использование контрактов',
                   'Объём транзакций', 'Объём бриджа', 'Количество поинтов', 'Дроп доступен']
        column_heights = [66.33, 21.22, 21, 24.11, 16.78, 13.33, 18.11, 13.56]

        for i in range(len(headers)):
            worksheet.write(0, i, headers[i], table_head)
            worksheet.set_column(i, i, column_heights[i])

        end = 0

        for i in range(len(data)):
            if not data[i]['status']:
                continue

            worksheet.merge_range(end + 1, 0, end + 3, 0, data[i]['data']['address'], border_format)

            worksheet.write(end + 1, 1, '3 месяца', success_format if 3 in data[i]['data']['criteria'][
                'transactions_over_time'] else unsuccess_format)
            worksheet.write(end + 2, 1, '6 месяцев', success_format if 6 in data[i]['data']['criteria'][
                'transactions_over_time'] else unsuccess_format)
            worksheet.write(end + 3, 1, '9 месяцев', success_format if 9 in data[i]['data']['criteria'][
                'transactions_over_time'] else unsuccess_format)

            worksheet.write(end + 1, 2, '25', success_format if 25 in data[i]['data']['criteria'][
                'transactions_frequency'] else unsuccess_format)
            worksheet.write(end + 2, 2, '50', success_format if 50 in data[i]['data']['criteria'][
                'transactions_frequency'] else unsuccess_format)
            worksheet.write(end + 3, 2, '100', success_format if 100 in data[i]['data']['criteria'][
                'transactions_frequency'] else unsuccess_format)

            worksheet.write(end + 1, 3, '10', success_format if 10 in data[i]['data']['criteria'][
                'contracts_variety'] else unsuccess_format)
            worksheet.write(end + 2, 3, '25', success_format if 25 in data[i]['data']['criteria'][
                'contracts_variety'] else unsuccess_format)
            worksheet.write(end + 3, 3, '50', success_format if 50 in data[i]['data']['criteria'][
                'contracts_variety'] else unsuccess_format)

            worksheet.write(end + 1, 4, '1000', success_format if 1000 in data[i]['data']['criteria'][
                'transaction_volume'] else unsuccess_format)
            worksheet.write(end + 2, 4, '5000', success_format if 5000 in data[i]['data']['criteria'][
                'transaction_volume'] else unsuccess_format)
            worksheet.write(end + 3, 4, '10000', success_format if 10000 in data[i]['data']['criteria'][
                'transaction_volume'] else unsuccess_format)

            bridge_volume = [str(_) for _ in data[i]['data']['criteria']['bridge_volume']]
            worksheet.merge_range(end + 1, 5, end + 3, 5, ', '.join(bridge_volume), border_format)

            worksheet.merge_range(end + 1, 6, end + 3, 6, data[i]['data']['points'], border_format)

            worksheet.merge_range(end + 1, 7, end + 3, 7, 'ДА' if data[i]['data']['eligible'] else 'НЕТ',
                                  success_format if data[i]['data']['eligible'] else 'НЕТ')

            end += 3

        if [i for i in data if not i['status']]:
            end += 1

            headers = ['Адрес', 'Причина ошибки']
            for i in range(len(headers)):
                worksheet.write(end + 1, i, headers[i], table_head)
            end += 1

            for i in range(len(data)):
                if data[i]['status']:
                    continue

                worksheet.write(end + 1, 0, data[i]['address'], border_format)
                worksheet.write(end + 1, 1, data[i]['data'], wrap_border_format)

                end += 1

        workbook.close()


def load_wallets():
    if path.exists('wallets.txt'):
        wallets = open('wallets.txt').read().split('\n')
        if not wallets[-1]:
            wallets = wallets[:-1]
        return wallets
    else:
        return FileNotFoundError('Файл wallets.txt не найден')


def init_exit():
    input('Нажмите Enter, чтобы выйти...')
    exit()


def main():
    wallets = load_wallets()
    if isinstance(wallets, FileNotFoundError):
        print(f'Не удалось загрузить кошельки: {wallets}\n')
        init_exit()

    print(f'Загружено {Fore.GREEN}{len(wallets)}{Fore.RESET} кошелька(ов)\n')

    check_result = check_wallets(wallets)

    print(f'{Fore.GREEN}\nПроверка завершена!{Fore.RESET}\n')

    user_input = input('Выберите тип отчёта (txt или xlsx): ')

    if user_input == 'txt':
        print(f'\nОтчёт будет выполнен в виде {Fore.GREEN}txt{Fore.RESET} файла\n')
    elif user_input == 'xlsx':
        print(f'\nОтчёт будет выполнен в виде {Fore.GREEN}xlsx{Fore.RESET} файла\n')
    else:
        user_input = 'xlsx'
        print(f'\nВвод не распознан. Отчёт будет выполнен в виде {Fore.GREEN}xlsx{Fore.RESET} файла\n')

    create_report(check_result, user_input)

    print(f'{Fore.GREEN}Отчёт сохранён!{Fore.RESET}\n')
    init_exit()


if __name__ == '__main__':
    init()
    main()
