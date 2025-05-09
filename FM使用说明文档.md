# 使用说明文档 - FileManipulator类

FileManipulator类是一个用于文件操作的工具类，它提供了复制文件夹、删除文件、重命名文件以及编辑docx文件的功能。下面是对每个方法的详细说明和使用示例。

## 1. 初始化FileManipulator对象

```python
file_manipulator = FileManipulator(str_oldpath, str_newpath, max_file_dict)
```

* `str_oldpath` (str): 源文件夹的路径。
* `str_newpath` (str): 目标文件夹的路径。
* `max_file_dict` (dict): 包含各类别中文件名最高数字索引的字典。

## 2. 复制文件夹

```python
file_manipulator.cp_files()
```

将源文件夹中的内容复制到目标文件夹中，并根据各类别中文件名的最高数字索引进行更新。如果目标文件夹已经存在，将创建一个带有时间戳的同名目录，并清空目录内容。

## 3. 删除文件

```python
file_manipulator.del_files()
```

删除目标路径下的文件。该方法会遍历目标路径下的所有子项，并删除以"~$"开头的临时文件和文件夹内的所有内容。

## 4. 重命名文件

```python
file_manipulator.ren_files()
```

对指定目录下的文件进行重命名。该方法会遍历目标路径下的所有子项，并检查文件名中是否包含括号。如果找到括号，将使用括号内的内容替换为文件夹的名称。

## 5. 编辑docx文件

```python
file_manipulator.edt_docx()
```

编辑指定目录下的docx文件。该方法会遍历目标路径下的所有子项，并根据特定条件对docx文件进行编辑。具体地，它会修改表头、高亮包含数字的单元格，并根据特定的表格结构进行高亮操作。

## 6. 打印文件树

```python
file_manipulator.print_directory_tree(path)
```

打印指定路径下的文件树。该方法会递归遍历指定路径下的文件和文件夹，并以树状结构进行打印。

## 7. 执行文件操作

```python
file_manipulator.execute_operations()
```

执行所有文件操作的方法。该方法会按照以下顺序依次执行：复制文件夹、删除文件、重命名文件、编辑docx文件，并最后打印目标文件夹的文件树。

以上是FileManipulator类的使用说明文档。通过实例化该类并调用其方法，您可以方便地进行文件操作，包括复制、删除、重命名和编辑docx文件。请根据您的需求和具体场景使用相应的方法。如有任何问题，请随时向我提问。