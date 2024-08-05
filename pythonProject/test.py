print(r'\(~_~)/ \(~_~)""\n/')
print('''Line 1
Line 2
Line 3''')
print('\"To be, or not to be\": that is the question.\nWhether it\'s nobler in the mind to suffer.')
print(r'''"To be, or not to be": that is the question.
Whether it's nobler in the mind to suffer.''')
print(f"ssdf'n'nnh")
name = "Alice"
print(f'He said, "Hello, {name}!"')
print(f"I'm {name}!")
print(f"List a contains:\n{name}")
file_path = "C:\\Users\\Public\\Documents\\file.txt"
print(f"The file is located at: {file_path}")
def greet(name):
    return f"Hello, {name.capitalize()}!"
user = "alice"
print(greet(user))
# 字符串模板
template = 'Hello {}'
# 模板数据内容
world = 'World'
result = template.format(world)
print(result) # ==> Hello World
# 指定{}的名字w,c,b,i
template = 'Hello {w}, Hello {c}, Hello {b}'
world = 'World'
china = 'China'
beijing = 'Beijing'
# 指定名字对应的模板数据内容
result = template.format(w = world, c = china, b = beijing)
print(result) # ==> Hello World, Hello China, Hello Beijing.