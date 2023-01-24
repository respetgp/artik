# import openpyxl as op
# import random
# import os
#
# # определяем файл
# file_name = 'my_artikel.xlsx'
#
# # Вводим переменные
# new_words = []
#
#
# def read_list_from_excel(file_exel: str) -> list:
#     '''создаем функцию по состалению из файла эксель списка'''
#
#     words_list_local = []
#     wb = op.load_workbook(file_exel, data_only=True)
#     sheet = wb.active
#     max_rows = sheet.max_row
#
#     for i in range(2, max_rows + 1):
#         article = sheet.cell(row=i, column=2).value
#         words = sheet.cell(row=i, column=3).value
#         translate = sheet.cell(row=i, column=4).value
#
#         if words is not None:
#             if len(words) != 0:
#                 words_list_local.append([article, words, translate])
#     return words_list_local
#
#
# def check_result_file(file_txt: str) -> list:
#     '''создаем фунцию проверки наличие файла для записи результата'''
#     new_words_local = []
#
#     if not os.path.exists(file_txt):
#         open(file_txt, 'tw', encoding='utf-8')
#     else:
#         with open(file_txt, 'r', encoding='utf-8') as f:
#             new_words_local = f.read().splitlines()
#
#     return new_words_local
#
# if __name__ == '__main__':
#
#     words_list_excel_2 = read_list_from_excel(file_name)
#     words_list_excel = []
#     new_words = check_result_file('new_words.txt')
#
#     # удаляем в списке для изучения слова из заученных
#     for i in range(0, len(words_list_excel_2)):
#         if not words_list_excel_2[i][1] in new_words:
#             words_list_excel.append(words_list_excel_2[i])
#         else:
#             print(words_list_excel_2[i][1])
#
#
#     count_words = int(input('введите количество слов: '))
#     count_true_anwords = 0
#
#     for count_words_index in range(0, count_words):
#         rand_words = random.randint(0, len(words_list_excel))
#         print(f'\n{words_list_excel[rand_words][1]}')
#         article_choice = input('выберете правильный артикль. напишите "der","die" или "das": ')
#         if article_choice == words_list_excel[rand_words][0]:  # это артикль
#             print(f'{" ".join(words_list_excel[rand_words])} \nправильно. вы большой молодец')
#             count_true_anwords += 1
#             new_words.append(words_list_excel[rand_words][1])
#
#         else:
#             print(f'{" ".join(words_list_excel[rand_words])} \nслово добавлено в список изучения')
#
#         with open("new_words.txt", "w", encoding="utf-8") as f:
#             for i in range(0, len(new_words)):
#                 f.write(new_words[i] + '\n')
#
#     print(f'''\nколичество правильных ответов {count_true_anwords} из {count_words}
#           Эффективность тренировки: {round((count_true_anwords / count_words * 100), 2)} процентов''')
#     print(f'''результат освоения вашего курса {len(new_words)} из {len(words_list_excel_2)}
#           Статистика освония курса: {round((len(new_words) / len(words_list_excel_2) * 100), 2)} процентов''')