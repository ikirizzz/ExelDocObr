
# исходные файла
filename1 = 'gentd.xlsx'
filename2 = 'emp.xlsx'
filename3 = 'spo.xlsx'
# промежуточные стандартизированные
res_filename1 = 'NewQentd.xlsx'
res_filename2 = 'NewEmp.xlsx'
res_filename3 = 'NewSpo.xlsx'
# промежуточный общий
cr_filename1 = 'CRes.xlsx'
# результирующий
pr_filename1 = 'PRIM12.xlsx'



#есть нагрузки (3 штуки) мы их подгружаем, чтобы стандартизировать (создаются новые файлы)
processing_file(filename1, res_filename1)
processing_file(filename2, res_filename2)
processing_file(filename3, res_filename3)

#создаем общий файл из предыдущих, он также обработывается и собирает все предметы кафедр группам
creat_file(res_filename1, res_filename2, res_filename3, cr_filename1)

# Вызовите функцию для копирования данных
copy_data_between_workbooks(cr_filename1, pr_filename1, '308', magistr=1, napravl= "Экономика")
copy_data_between_workbooks(cr_filename1, pr_filename1, '508', magistr=1, napravl= "Экономика")
copy_data_between_workbooks(cr_filename1, pr_filename1, '129', napravl= "Строительство")
copy_data_between_workbooks(cr_filename1, pr_filename1, '229', napravl= "Строительство")
copy_data_between_workbooks(cr_filename1, pr_filename1, '329', napravl= "Строительство")
copy_data_between_workbooks(cr_filename1, pr_filename1, '429', napravl= "Строительство")
