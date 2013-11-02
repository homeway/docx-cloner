#language: zh-CN

功能: 读Docx内标签定义
  这里要确认标签读取的正确性，然后再进入替换阶段
  1、主要解决的问题包括：将docx文件拆包、找到对应的文件位置
  2、xml标记可能是散开的，例如"{name}"在docx文件内部表示中，"{"、"name"、"}"是各自独立的xml标记
  3、替换逻辑，希望使用DSL在程序中指定，因此不应该限定到底使用"{name}"还是"$name$"做标签标识

  背景: 可读的示例文件列举
    假如"docx-examples"示例文件夹中存在一个"read-single-tags.docx"的文件

  场景大纲: 简单地读取词语替换标签
    这是最简单的情形，例如将标签{name}，替换为真正的姓名。

    那么程序应该能读到"<tagname>"这个标签词

    例子: 读取标签的例子
      "{}"可作为默认的正则表达式设计，在DSL中无需指定
      程序应该支持中文（以及其它UTF8字符）

      | tagname |
      | {name}  |
      | {Name}  |
      | {NAME}  |
      | {{名字}} |
      | $名字$   |

  场景: 所读取的替换标签应与docx内的xml判断一致
    这是一个深入到docx的内部结构的场景
    假如文件内包含"{name}"标签
    那么应该解析这个标签应该能读到这样的XML片段：
    """
    <w:r w:rsidR="000F595B">
      <w:rPr>
        <w:rFonts w:hint="eastAsia"/>
        <w:lang w:eastAsia="zh-CN"/>
      </w:rPr>
      <w:t>{</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:hint="eastAsia"/>
        <w:lang w:eastAsia="zh-CN"/>
      </w:rPr>
      <w:t>n</w:t>
    </w:r>
    <w:r w:rsidR="000F595B">
      <w:rPr>
        <w:rFonts w:hint="eastAsia"/>
        <w:lang w:eastAsia="zh-CN"/>
      </w:rPr>
      <w:t>ame}</w:t>
    </w:r>
    """

  场景: 读取表格行替换标签
    这通常是在表格上追加行所使用的

  场景: 读取文档信息标签
    包括标题、摘要、作者、邮件等设置信息

  场景: 读取图像标签
    这是做图像替换时使用的
