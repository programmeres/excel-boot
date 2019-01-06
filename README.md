# Easy-POI

Easy-POI是一款Excel导入导出解决方案组成的轻量级开源组件。

**如果喜欢或愿意使用, 请star并且Watch本项目**

**如果是企业使用, 请通过修改本文件的企业列表告诉我企业名称(非盈利用途).**

发现解决bug并已自测,pullRequest后,可以通过邮件告知我们（magic_app@126.com）, 第一时间合并并且发布最新版本
## 使用企业列表：

## 功能简介
1. 浏览器导出Excel文件（支持单/多sheet）

2. 浏览器导出Excel模板文件  

3. 指定路径生成Excel文件（支持单/多sheet）

4. 返回Excel文件（支持单/多sheet）的OutputStream, 一般用于将Excel文件上传到远程, 例如FTP

5. 导入Excel文件（支持单/多sheet）

## 解决的问题
1.解决导出大量数据造成的内存溢出问题（支持分页查询数据库、采用poi官方推荐api（SXSSFWorkbook）, 实现指定行数刷新到磁盘）

2.解决导入大量数据造成的内存溢出问题（分页插入数据库、采用poi官方推荐api（XSSF and SAX），采用SAX模式一行行读取到内存当中去)

3.解决含有占位符的空假行造成的读空值问题

4.解决Long类型或者BigDecimal的精度不准问题

## 组件特色
1.导入可以自定义解析成功或失败的处理逻辑

2.导出支持分页查询、全量查询, 自定义每条数据的处理逻辑

3.内置缓存,3万条11列数据,第一次导出2.2s左右、第二次导出在1.4s左右；第一次导入3.5s左右、第二次导入2.5s左右

4.注解操作, 轻量且便捷

5.内置常用正则表达式类RegexConst（身份证号、手机号、金额、邮件）

6.适配单元格宽度(单元格内容最长不得超过20个汉字)

7.假如出现异常,Sheet、行、列位置也都一并打印

8.注解中的用户自定义字符串信息以及Excel信息已全部trim,不用担心存在前后空格的风险

9.Excel样式简洁、大方、美观

## 组件需知
导入和导出只支持尾缀为xlsx的Excel文件、标注注解的属性顺序即Excel列的排列顺序、时间转化格式(dateFormat)默认为“yyyy-MM-dd HH:mm:ss“.

### 导入
1.当导入Excel, 读取到空行, 则停止读取当前Sheet的后面数据行

2.导入Excel文件, 单元格格式使用文本或者常规, 防止出现不可预测异常

3.导入字段类型支持:Date、Short（short）、Integer（int）、Double（double）、Long（long）、Float（float）、BigDecimal、String类型

4.导入BigDecimal字段精度默认为2, roundingMode默认为BigDecimal.ROUND_HALF_EVEN

5.第一行有效单元格内必须包含内容并且以第一行为依据, 导入Excel文件列数必须等于标注注解的属性数量

6.Date类型字段,Excel与时间转化格式(dateFormat)相比,格式要保持一致（反例：2018/12/31和“yyyy-MM-dd“）并且长度要一致或更长（反例："2018-12-31"和yyyy-MM-dd HH:mm:ss"）,否则SimpleDateFormat将解析失败,报 “Unparseable date:”

### 导出
1.导出BigDecimal字段默认不进行精度格式化

2.分页查询默认从第一页开始, 每页3000条

3.Excel每超过2000条数据, 将内存中的数据刷新到磁盘当中去

4.使用分Sheet导出方法, 每8万行数据分Sheet

5.当使用（exportResponse、exportStream、generateExcelStream）方法时, 当单个Sheet超过100万条则会分Sheet

6.标注属性类型要与数据库类型保持一致


## 扩展
继承EasyPoi类, 可以使用子类构造器覆盖以下默认参数
```java
    //Excel自动刷新到磁盘的数量
    public static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 2000;
    //分页条数
    public static final int DEFAULT_PAGE_SIZE = 3000;
    //分Sheet条数
    public static final int DEFAULT_RECORD_COUNT_PEER_SHEET = 80000;
```
## 版本
当前为1.0.0版本, 2.0.0版正在开发当中

## 使用手册
1.引入Maven依赖

2.将需要导出或者导入的实体属性上标注@ExportField或@ImportField注解

3.直接调用导出或导入API即可

### POM.xml 

```xml
<dependency>
	<groupId>io.github.magic-core</groupId>
	<artifactId>easy-poi</artifactId>
	<version>2.0.0</version>
</dependency>
```
### @ExportField

```java
/**
* 导出注解功能介绍
*/
public @interface ExportField {

    /**
     * excel列名称
     */
    String columnName();

    /**
     * 默认单元格值
     */
    String defaultCellValue() default "";

    /**
     * 日期格式 默认 yyyy-MM-dd HH:mm:ss
     */
    String dateFormat() default "yyyy-MM-dd HH:mm:ss";

    /**
     * BigDecimal精度 默认:-1(默认不开启BigDecimal格式化)
     */
    int scale() default -1;

    /**
     * BigDecimal 舍入规则 默认:BigDecimal.ROUND_HALF_EVEN
     */
    int roundingMode() default BigDecimal.ROUND_HALF_EVEN;
}
```
```java
/**
* 导出注解Demo
*/
public class ExportFielddemo {
@ExportField(columnName = "ID", defaultCellValue = "1")
    private Integer id;

@ExportField(columnName = "姓名", defaultCellValue = "张三")
    private String name;
    
@ExportField(columnName = "收入金额", defaultCellValue = "100", scale = 2, roundingMode=BigDecimal.ROUND_HALF_EVEN)
    private BigDecimal money;

@ExportField(columnName = "创建时间", dateFormat="yyyy-MM-dd", defaultCellValue = "2019-01-01")
    private Date createTime;
}
```

### @ImportField

```java
/**
* 导入注解功能介绍
*/
public @interface ImportField {

    /**
     * @return 是否必填
     */
    boolean required() default false;

    /**
     * 日期格式 默认 yyyy-MM-dd HH:mm:ss
     */
    String dateFormat() default "yyyy-MM-dd HH:mm:ss";

    /**
     * 正则表达式校验
     */
    String regex() default "";

    /**
     * 正则表达式校验失败返回的错误信息, regex配置后生效
     */
    String regexMessage() default "正则表达式验证失败";

    /**
     * BigDecimal精度 默认:2
     */
    int scale() default 2;

    /**
     * BigDecimal 舍入规则 默认:BigDecimal.ROUND_HALF_EVEN
     */
    int roundingMode() default BigDecimal.ROUND_HALF_EVEN;
}
```
```java
/**
* 导入注解Demo
*/
public class ImportField {
@ImportField(required = true)
    private Integer id;

@ImportField(regex = IDCARD_REGEX, regexMessage="身份证校验失败")
    private String idCard;
    
@ImportField(scale = 2, roundingMode=BigDecimal.ROUND_HALF_EVEN)
    private BigDecimal money;

@ImportField(dateFormat="yyyy-MM-dd")
    private Date createTime;
}
```
### 导出Demo
```java
/**
 * 导出Demo
 */
public class ExportDemo {
    /**
     * 浏览器导出Excel
     *
     * @param httpServletResponse
     */
    public void exportResponse(HttpServletResponse httpServletResponse) {
        ParamEntity queryQaram = new ParamEntity();
        EasyPoi.ExportBuilder(httpServletResponse, "Excel文件名", AnnotationEntity.class).exportResponse(queryQaram,
                new ExportFunction<ParamEntity, ResultEntity>() {
                    /**
                     * @param queryQaram 查询条件对象
                     * @param pageNum    当前页数,从1开始
                     * @param pageSize   每页条数,默认3000
                     * @return
                     */
                    @Override
                    public List<ResultEntity> pageQuery(ParamEntity queryQaram, int pageNum, int pageSize) {

                        //分页查询操作

                        return new ArrayList<ResultEntity>();
                    }

                    /**
                     * 将查询出来的每条数据进行转换
                     *
                     * @param o
                     */
                    @Override
                    public AnnotationEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 浏览器多sheet导出Excel
     *
     * @param httpServletResponse
     */
    public void exportMultiSheetResponse(HttpServletResponse httpServletResponse) {
        ParamEntity queryQaram = new ParamEntity();
        EasyPoi.ExportBuilder(httpServletResponse, "Excel文件名", AnnotationEntity.class).exportMultiSheetStream(queryQaram,
                new ExportFunction<ParamEntity, ResultEntity>() {
                    /**
                     * @param queryQaram 查询条件对象
                     * @param pageNum    当前页数,从1开始
                     * @param pageSize   每页条数,默认3000
                     * @return
                     */
                    @Override
                    public List<ResultEntity> pageQuery(ParamEntity queryQaram, int pageNum, int pageSize) {

                        //分页查询操作

                        return new ArrayList<ResultEntity>();
                    }

                    /**
                     * 将查询出来的每条数据进行转换
                     *
                     * @param o
                     */
                    @Override
                    public AnnotationEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 导出Excel到指定路径
     */
    public void exportStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        EasyPoi.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", AnnotationEntity.class)
                .exportStream(queryQaram, new ExportFunction<ParamEntity, ResultEntity>() {
                    /**
                     * @param queryQaram 查询条件对象
                     * @param pageNum    当前页数,从1开始
                     * @param pageSize   每页条数,默认3000
                     * @return
                     */
                    @Override
                    public List<ResultEntity> pageQuery(ParamEntity queryQaram, int pageNum, int pageSize) {

                        //分页查询操作

                        return new ArrayList<ResultEntity>();
                    }

                    /**
                     * 将查询出来的每条数据进行转换
                     *
                     * @param o
                     */
                    @Override
                    public ResultEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 导出多sheet Excel到指定路径
     */
    @RequestMapping(value = "exportResponse")
    public void exportMultiSheetStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        EasyPoi.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", AnnotationEntity.class)
                .exportMultiSheetStream(queryQaram, new ExportFunction<ParamEntity, ResultEntity>() {
                    /**
                     * @param queryQaram 查询条件对象
                     * @param pageNum    当前页数,从1开始
                     * @param pageSize   每页条数,默认3000
                     * @return
                     */
                    @Override
                    public List<ResultEntity> pageQuery(ParamEntity queryQaram, int pageNum, int pageSize) {

                        //分页查询操作

                        return new ArrayList<ResultEntity>();
                    }

                    /**
                     * 将查询出来的每条数据进行转换
                     *
                     * @param o
                     */
                    @Override
                    public AnnotationEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 生成Excel OutputStream对象
     */
    public void generateExcelStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        OutputStream outputStream = EasyPoi.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", AnnotationEntity.class)
                .generateExcelStream(queryQaram, new ExportFunction<ParamEntity, ResultEntity>() {
                    /**
                     * @param queryQaram 查询条件对象
                     * @param pageNum    当前页数,从1开始
                     * @param pageSize   每页条数,默认3000
                     * @return
                     */
                    @Override
                    public List<ResultEntity> pageQuery(ParamEntity queryQaram, int pageNum, int pageSize) {

                        //分页查询操作

                        return new ArrayList<ResultEntity>();
                    }

                    /**
                     * 将查询出来的每条数据进行转换
                     *
                     * @param o
                     */
                    @Override
                    public AnnotationEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 生成多Sheet Excel OutputStream对象
     */
    public void generateMultiSheetExcelStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        OutputStream outputStream = EasyPoi.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", AnnotationEntity.class)
                .generateMultiSheetExcelStream(queryQaram, new ExportFunction<ParamEntity, ResultEntity>() {
                    /**
                     * @param queryQaram 查询条件对象
                     * @param pageNum    当前页数,从1开始
                     * @param pageSize   每页条数,默认3000
                     * @return
                     */
                    @Override
                    public List<ResultEntity> pageQuery(ParamEntity queryQaram, int pageNum, int pageSize) {

                        //分页查询操作

                        return new ArrayList<ResultEntity>();
                    }

                    /**
                     * 将查询出来的每条数据进行转换
                     *
                     * @param o
                     */
                    @Override
                    public AnnotationEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 导出Excel模板
     */
    public void exportTemplate(HttpServletResponse httpServletResponse) {
        EasyPoi.ExportBuilder(httpServletResponse, "Excel模板名称", AnnotationEntity.class).exportTemplate();
    }
}
```
### 导入Demo
```java
/**
 * 导入Demo
 */
public class ImportDemo {
    public void importExcel() throws IOException {
        EasyPoi.ImportBuilder(new FileInputStream(new File("C:\\Users\\导入Excel文件.xlsx")),  AnnotationEntity.class)
                .importExcel(new ExcelImportFunction<AnnotationEntity>() {

                    /**
                     * @param sheetIndex 当前执行的Sheet的索引, 从1开始
                     * @param rowIndex 当前执行的行数, 从1开始
                     * @param resultEntity Excel行数据的实体
                     */
                    @Override
                    public void onProcess(int sheetIndex,  int rowIndex,  AnnotationEntity resultEntity) {
                        //对每条数据自定义校验以及操作
                        //分页插入：当读取行数到达用户自定义条数执行插入数据库操作
                    }

                    /**
                     * @param errorEntity 错误信息实体
                     */
                    @Override
                    public void onError(ErrorEntity errorEntity) {
                        //操作每条数据非空和正则校验后的错误信息
                    }
                });
    }
}
```