def str_count(text, substr):    #количество вхождений подстроки в строку
    return len(text.split(substr))-1


text='БАСО-01-19 10.05.03(КБ-1)'
substr='-'
print(str_count(text, substr))