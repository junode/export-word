## 总结：
不与考虑的工具：Jacob、JSP、iText、java2Word、PageOffice、Spire.DOC、jasperreports；予以考虑的工具：Apache POI、FreeMarker

+ Jacob不考虑，因为其服务器只能是windows平台，不支持unix和linux，且服务器上必须安装微软Office。
+ iText，只能生成rtf格式的文档，这个是致命的，对于客户体验来说，这个是不可接受的。
+ java2Word，需要windows支持，ruo-yi应该更加通用，不因仅word导出就让ruo-yi应用环境受限。
+ PageOffice，不能在服务器端生成文件，ruo-yi应该更加通用，不因仅word导出就让ruo-yi应用环境受限。
+ Spire.DOC：收费的。
+ jasperreports：在maven 仓库中 引用数过低【更实际的原因是我拉不下它的pom依赖。】，还有原因就是我看到它依赖的包有点多，让我觉得它有点重了，怕对项目造成不良影响。

## java导出word的若干方式及比较

> https://www.cnblogs.com/ziwuxian/p/8985678.html
> https://www.cnblogs.com/yuluoxingkong/p/9298270.html

1：Jacob是Java-COM Bridge的缩写，它在Java与微软的COM组件之间构建一座桥梁。通过Jacob实现了在Java平台上对微软Office的COM接口进行调用。

- 优点：调用微软Office的COM接口，生成的word文件格式规范。

- 缺点：服务器只能是windows平台，不支持unix和linux，且服务器上必须安装微软Office。

2：Apache POI包括一系列的API，它们可以操作基于MicroSoft OLE 2 Compound Document Format的各种格式文件，可以通过这些API在Java中读写Excel、Word等文件。

- 优点：跨平台支持windows、unix和linux。
- 缺点：相对与对word文件的处理来说，POI更适合excel处理，对于word实现一些简单文件的操作凑合，不能设置样式且生成的word文件格式不够规范。

3：Java2word是一个在java程序中调用 MS Office Word 文档的组件(类库)。该组件提供了一组简单的接口，以便java程序调用他的服务操作Word 文档。 这些服务包括： 打开文档、新建文档、查找文字、替换文字，插入文字、插入图片、插入表格，在书签处插入文字、插入图片、插入表格等。

- 优点：足够简单，操作起来要比FreeMarker简单的多。
- 缺点：没有FreeMarker强大，不能够根据模版生成Word文档，word的文档的样式等信息都不能够很好的操作。

4：FreeMarker生成word文档的功能是由XML+FreeMarker来实现的。先把word文件另存为xml，在xml文件中插入特殊的字符串占位符，将xml翻译为FreeMarker模板，最后用java来解析FreeMarker模板，编码调用FreeMarker实现文本替换并输出Doc。

- 优点：比Java2word功能强大，也是纯Java编程。
- 缺点：生成的文件本质上是xml，不是真正的word文件格式，有很多常用的word格式无法处理或表现怪异，比如：超链、换行、乱码、部分生成的文件打不开等。

5：PageOffice生成word文件。PageOffice封装了微软Office繁琐的vba接口，提供了简洁易用的Java编程对象，支持生成word文件，同时实现了在线编辑word文档和读取word文档内容。

- 优点：跨平台支持windows、unix和linux，生成word文件格式标准，支持文本、图片、表格、字体、段落、颜色、超链、页眉等各种格式的操作，支持多word合并，无需处理并发，不耗费服务器资源，运行稳定。

- 缺点：必须在客户端生成文件（可以不显示界面），不支持纯服务器端生成文件。


| 所用技术    | 优点                                                         | 缺点                                                         |
| ----------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| Jacob       | 功能强大                                                     | 代码量大，设置样式繁琐；需要windows平台支持，无法跨平台      |
| Apache POI  | 读写excel功能强大、操作简单                                  | 一般只用它读取word，能够创建简单的word，不能设置样式，功能太少 |
| Java2word   | 功能强大，操作简单                                           | 能满足一般要求，不支持07格式，国人开发的，参考资料较多，需要windows平台支持 |
| iText       | 功能全，能满足一般要求                                       | 不能直接生成或操作doc文档，只能生成rtf格式的文档，rtf也可以用word打开 |
| JSP         | 操作简单，代码量少                                           | 能把当前页面导出简单的word，不能设置样式，美观性差，无法操作word |
| XML（最佳） | 代码量少，样式、内容容易控制，打印不变形，完全符合office标准 | 需要提前设计好word模板，把需要替换的地方用特殊标记标出来     |
| Spire.DOC   | 提供了很多Demo方便学习、兼容各word版本                       | 非开源、收费                                                 |

## Word导出个人整理
- Apache POI
  - demo
    - https://www.baeldung.com/java-microsoft-word-with-apache-poi
    - https://www.swtestacademy.com/generate-word-document-with-java/
  - 基于POI的封装：Poi-tl Documentation
    - 官方定义：poi-tl（poi template language）是Word模板引擎，使用Word模板和数据创建很棒的Word文档。
    - github地址：https://github.com/Sayi/poi-tl
    - 中文文档链接：http://deepoove.com/poi-tl/
- freemarker
  - demo：https://github.com/chenhongen/Freemarker2Word
- Java2Word
  - github地址：https://github.com/leonardoanalista/java2word
  - demo：https://www.jianshu.com/p/f773889bf627?utm_campaign=maleskine&utm_content=note&utm_medium=seo_notes&utm_source=recommendation
- jasper
  - demo：https://github.com/javahowtos/jasper-export-demo
- aspose-words
  - Generate Word Documents from Templates Dynamically using Java：https://blog.aspose.com/2020/01/14/generate-word-documents-from-templates-dynamically-using-java/
- Spire.DOC
  - 由e-iceblue出品，该工具没有找到maven的相关依赖，仅有jar包，不过看下面demo，非常容易上手，应该是付费的。
  - 官方示例：https://www.e-iceblue.com/Tutorials/Java/Spire.Doc-for-Java/Program-Guide/Document-Operation/Create-Word-Document-in-Java.html
- docx4j-ImportXHTML
  - github地址：https://github.com/plutext/docx4j-ImportXHTML


## 当前项目中已有demo说明

### ContentToWordController.class
> 依赖Apache POI，将获取到的String转为word存储，无样式，只是单纯的导出String为word存储，
> 若有需要应该可以添加插入图片。

### TextAndPictureToWordDemo.class
> Apache POI包括一系列的API，它们可以操作基于MicroSoft OLE 2 Compound Document Format的各种格式文件，可以通过这些API在Java中读写Excel、Word等文件。
- 优点：跨平台支持windows、unix和linux。
- 缺点：相对与对word文件的处理来说，POI更适合excel处理，对于word实现一些简单文件的操作凑合，不能设置样式且生成的word文件格式不够规范。

### POITLController.class
> 对POI进行封装，提供非单易用的API，可以用其做简历。<br />

衡量点：<br />
+ 1、版本稳定性；
+ 2、对于仅需导出word，若无复杂样式需求，我觉得用该工具过重；
+ 3、与ruo-yi 是否兼容

### freemarker.Demo.class
> 采用Freemarker模板生成word是由XML+FreeMarker实现的。 XML文件定义模板，然后Freemarker解析模板输出到word。
+ 优点：比Java2word功能强大，也是纯Java编程。
+ 缺点：生成的文件本质上是xml，不是真正的word文件格式，有很多常用的word格式无法处理或表现怪异，比如：超链、换行、乱码、部分生成的文件打不开等。

个人考量点：
+ XML模板文件的复杂度。对于我们当前不复杂的功能，我觉得可以使用freemarker导出word。

### JavaToWordController.class
Java2word是一个在java程序中调用 MS Office Word 文档的组件(类库)。该组件提供了一组简单的接口，以便java程序调用他的服务操作Word 文档。 这些服务包括： 打开文档、新建文档、查找文字、替换文字，插入文字、插入图片、插入表格，在书签处插入文字、插入图片、插入表格等。

- 优点：足够简单，操作起来要比FreeMarker简单的多。
- 缺点：没有FreeMarker强大，不能够根据模版生成Word文档，word的文档的样式等信息都不能够很好的操作。
  个人考量点：
  javaToWord只有jar包，没有pom依赖。当然其代码量相当少，可以直接引入做二次开发，但其需要windows平台支持。

