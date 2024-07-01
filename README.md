# mapping_extractor
Extracts mappings from KUMA's normalizers and writes them to .xlsx file

# Описание

Данный скрипт позволяет из ресурса, экспортированного из KUMA с помощью веб-интерфейса, извлечь все нормализаторы и записать сопоставление полей нормализаторов, а также обогащение в .xlsx файл.

За расшифровку ресурса и приведение его к JSON отвечает kuma_package.py. За это отдельное спасибо [Morpheme777](https://github.com/Morpheme777). Оригинал этого скрипта можно найти [тут](https://github.com/Morpheme777/kuma_packages_encryption).

# Требования

python3.6+
- os
- json
- bson

```pip install bson```

- codecs
- Crypto
  
```pip install pycryptodome```

- openpyxl
- argparse

# Совместимость

Скрипт проверен на ресурсах KUMA 2.1.1.73 и KUMA 3.2.0.305

# Описание работы

Скрипт на вход получает зашифрованный файл с ресурсами KUMA, а также пароль от данного файла. Для каждого нормализатора в своей рабочей директории скрипт генерирует файл с именем `_<имя нормализатора>.xlsx`, который содержит маппинг и обогащение соответствующего нормализатора.

# Установка

```git clone https://github.com/KUMA-Community/mapping_extractor```

Либо просто скачайте файлы `mapping_extractor.py` и `kuma_package.py`.

# Использование

```
usage: mapping_extractor.py [-h] -f F -p P

options:
  -h, --help  show this help message and exit
  -f F        file with resources
  -p P        password for resources
```

# Известные ограничения

Для некоторых версий Excel при открытии файла .xlsx, сгенерированного данным скриптом может возникать ошибка

![image](https://github.com/koalapower/mapping_extractor/assets/28699284/b7e3158c-da3c-44d5-8f98-1f316c119b20)

В этом случае следует нажать на кнопку **Yes** (**Да**) и пересохранить файл. Тогда в дальнейшем данная ошибка возникать не будет.

P.S. если знаете, как пофиксить данную проблему - свяжитесь со мной =)
