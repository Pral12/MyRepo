import os
import platform
import sys

cwd = os.getcwd()
print("Current Working Directory:", cwd)
files = os.listdir('.')
print("Files and directories:", files)
home_dir = os.getenv('HOME')
os.environ['MY_VAR'] = 'my_value'
print("Home Directory:", home_dir)

os_name = platform.system()
print("Operating System:", os_name)
node_name = platform.node()
print("Node Name:", node_name)
os_release = platform.release()
print("OS Release:", os_release)
architecture = platform.architecture()
print("Architecture:", architecture)
processor = platform.processor()
print("Processor:", processor)
python_version = platform.python_version()
print("Python Version:", python_version)
python_compiler = platform.python_compiler()
print("Python Compiler:", python_compiler)

print("Operating System:", platform.system())
print("Node Name:", platform.node())
print("OS Release:", platform.release())
print("OS Version:", platform.version())
print("Machine:", platform.machine())
print("Processor:", platform.processor())
print("Architecture:", platform.architecture())
print("Python Version:", platform.python_version())
print("Python Compiler:", platform.python_compiler())


#if len(sys.argv) < 2:
#    print("Ошибка: недостаточно аргументов")
 #   sys.exit(1)

#print("Все аргументы указаны корректно")
#sys.exit(0)

print("Пути поиска модулей:")
for path in sys.path:
    print(path)

# Добавление нового пути
sys.path.append('/path/to/my/modules')
print("Обновленный список путей поиска модулей:", sys.path)

# Версия Python
print("Версия Python:", sys.version)

# Информация о платформе
print("Платформа:", sys.platform)

# Размер числа в байтах
print("Размер int:", sys.getsizeof(0), "байт")

# Список загруженных модулей
print("Загруженные модули:")
for module in sys.modules:
    print(module)