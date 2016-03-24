# coding: UTF-8
#version 1.0
import xlrd
import ArabianUnicodeMap
#input output file name
file = 'arabian.xls'
OUT = 'string.txt'
#空字符 与空白符
STR_NULL = ''
STR_SPACE = ' '
CANT_FIND = ord('#')
ENGLISH_INDEX_RANGE = list(range(0,123))
#file handle
def open_excel(file):
	with xlrd.open_workbook(file) as book:
		return book

def open_excel_by_index(file,index):
	book = open_excel(file)
	sheet = book.sheet_by_index(0)
	return sheet
    
def save_file(file="string.txt",string=''):
	with open(file,'w',encoding='utf-8') as text:
		text.write(string)
        #text.close()

#string handle 
def wash_space(string):
    if string == STR_NULL:
        string = STR_SPACE
    else:
        string.strip()
    
    return string

def arabian_unicode_map(list_letter):
    len_word = len(list_letter)
    list_unicode = []
    if len_word == 1:
        font_list = ArabianUnicodeMap.ARABIAN_FONT_DICT[list_letter[0]]
        list_unicode.append(font_list[ArabianUnicodeMap.ISOLATED]) 
    else:
        font_list = ArabianUnicodeMap.ARABIAN_FONT_DICT[list_letter[0]]
        list_unicode.append(font_list[ArabianUnicodeMap.INITIAL]) 
        for i in range(1,len_word-1):
            if list_letter[i] in ArabianUnicodeMap.ARABIAN_UNICODE_RANGE:
                font_list = ArabianUnicodeMap.ARABIAN_FONT_DICT[list_letter[i]]
                if list_letter[i-1] == 0x644 and list_letter[i] in ArabianUnicodeMap.ARIBIAN_UNICODE_WITH_0X644:
                    print('with 0x644')
                    font_list = ArabianUnicodeMap.ARABIAN_FONT_WITH_0X644_DICT[list_letter[i]]
                    list_unicode[len(list_unicode)-1] = font_list[ArabianUnicodeMap.MEDIAL]
                    #list_unicode.append(font_list[ArabianUnicodeMap.INITIAL])                
                elif list_letter[i-1] in ArabianUnicodeMap.ARIBIAN_UNICODE_SPECIAL or list_unicode[len(list_unicode)-1] in ArabianUnicodeMap.ARIBIAN_UNICODE_WITH_0X644_RANGE:
                    list_unicode.append(font_list[ArabianUnicodeMap.INITIAL])
                else:
                    list_unicode.append(font_list[ArabianUnicodeMap.MEDIAL])
            else:
                print('CANT_FIND',hex(list_letter[i]))
                list_unicode.append(CANT_FIND)                 
        if list_letter[len_word-1] in ArabianUnicodeMap.ARABIAN_UNICODE_RANGE:
            font_list = ArabianUnicodeMap.ARABIAN_FONT_DICT[list_letter[len_word-1]]
            if list_letter[len_word-2] in ArabianUnicodeMap.ARIBIAN_UNICODE_SPECIAL:
                list_unicode.append(font_list[ArabianUnicodeMap.ISOLATED])
            else:
                list_unicode.append(font_list[ArabianUnicodeMap.FINAL])
        else:
            list_unicode.append(CANT_FIND)
    return list_unicode

def arabian_unicode_to_index(list_unicode):
    for i in range(len(list_unicode)):
        if list_unicode[i] in ArabianUnicodeMap.ARABIAN_UNICODE_FORMS_B:
            list_unicode[i] = list_unicode[i] - 0xFE80 + 123
    return list_unicode

def word_to_unicode(str_word):
    list_letter = []
    for ch in str_word:
        list_letter.append(ord(ch))
    #print('list_letter',list_letter)
    return list_letter

def unicode_to_index(list_unicode):
    list_index = []
    if list_unicode[0] in ArabianUnicodeMap.ENGLISH_UNICODE_RANGE:
        list_index = list_unicode
    elif list_unicode[0] in ArabianUnicodeMap.ARABIAN_UNICODE_RANGE:
        list_unicode = arabian_unicode_map(list_unicode)
        list_index = arabian_unicode_to_index(list_unicode) 
    #print('list_index',list_unicode)
    return list_index
#返回一个词的list：index
def word_to_index(str_word):
    list_unicode_letter = word_to_unicode(str_word)
    list_index = unicode_to_index(list_unicode_letter)
    if  list_index[0] in ENGLISH_INDEX_RANGE:
        pass
    else:
        list_index.reverse()
    return list_index

#二维列表转字符串
def index_to_string_list(list_sentence_str_index):
    list_sentence_str = []

    for l in list_sentence_str_index:
        for i in range(len(l)):
            list_sentence_str.append(str(l[i]))
        if l != list_sentence_str_index[len(list_sentence_str_index)-1]:
            list_sentence_str.append(str(ord(STR_SPACE)))
    return list_sentence_str

def format_string_list(list_sentence_str_str):
    len_sentence = len(list_sentence_str_str)
    CONNECT_SYMBOL = ','
    str_buffer = CONNECT_SYMBOL.join(list_sentence_str_str)
    format_string = '{' + str(len_sentence) + ',' + str_buffer + '};'
    #format_string.replace('35','\'#\'')
    return format_string

def string_to_index(sentence):
    list_sentence_str_index = []
    sentence = wash_space(sentence)
    if len(sentence) == 0:
        sentence = STR_SPACE
    list_wordstr = sentence.split()
    if len(list_wordstr) == 0:
        list_wordstr.append(STR_SPACE)
    for str_word in list_wordstr:
        list_sentence_str_index.append(word_to_index(str_word))
    list_sentence_str_index.reverse()
    #print('list_sentence_str_index',list_sentence_str_index)
    list_sentence_str_str = index_to_string_list(list_sentence_str_index)
    format_string = format_string_list(list_sentence_str_str)
    return format_string            
   
data = STR_NULL
sheet = open_excel_by_index(file,0)
str_cloum = sheet.col_values(0)


for str_row in str_cloum:
	data = data + string_to_index(str_row) + '\n'




letter_collect = []
letter_count = 0		
for str_line in str_cloum:
    str_line = wash_space(str_line)
    for ch in str_line:
        if ch not in letter_collect:
            letter_collect.append(ch)
    
letter_collect = sorted(letter_collect)
for i in letter_collect:
	print(hex(ord(i)))
letter_num = len(letter_collect)
print("lenth:",letter_num)		
	
save_file(OUT,data)
print('done')
