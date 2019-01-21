<p align="center" id="e-b">
<img src="https://gitee.com/magicCore/codes/fb2jiwvrqlkcsotynpmdx78/raw?blob_name=excel-boot.png" >
    <p align="center">
        <a href="https://github.com/programmeres/easy-poi/releases">
            <img src="https://img.shields.io/github/release/programmeres/easy-poi.svg" >
        </a>
        <a href="https://opensource.org/licenses/artistic-license-2.0.php">
            <img src="https://img.shields.io/github/license/programmeres/easy-poi.svg" >
        </a>
        <a href="#e-b">
            <img src="https://img.shields.io/badge/coverage-100%25-red.svg" >
        </a>
        <a href="https://app.codacy.com/app/programmeres/easy-poi?utm_source=github.com&utm_medium=referral&utm_content=programmeres/easy-poi&utm_campaign=Badge_Grade_Dashboard">
            <img src="https://api.codacy.com/project/badge/Grade/6847cc8aa4154bee892b250a9bc846c9" >
        </a>
         <a href="https://gitee.com/nw1992/easy-poi#project-donate-overview">
            <img src="https://img.shields.io/badge/donate-%EF%BF%A5-orange.svg" >
        </a>
    </p>    
</p>

**Excel-Boot是一款Excel导入导出解决方案组成的轻量级开源组件。**

**如果喜欢或愿意使用, 请star本项目或者点击donate图标捐赠我们**

**如果是企业使用, 为了产品推广, 请通过评论、Issue、PullRequest README的合作企业告诉我们企业名称**

**如果有功能需求, 请修改页面底部的<建议功能投票列表>进行投票**

**如果有任何问题可以通过issue或者评论或者添加QQ群（716145748）告知我们, 尽力第一时间解决您的问题**
## 合作企业：

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/3a1ec066e9a542f682f61309ea10d820)](https://app.codacy.com/app/programmeres/excel-boot?utm_source=github.com&utm_medium=referral&utm_content=programmeres/excel-boot&utm_campaign=Badge_Grade_Dashboard)

## 开源库地址（同步更新）：

GitHub：<https://github.com/programmeres/excel-boot>

码云：<https://gitee.com/nw1992/easy-poi>

## 功能简介
1. 浏览器导出Excel文件（支持单/多sheet）

2. 浏览器导出Excel模板文件  

3. 指定路径生成Excel文件（支持单/多sheet）

4. 返回Excel文件（支持单/多sheet）的OutputStream, 一般用于将Excel文件上传到远程, 例如FTP

5. 导入Excel文件（支持单/多sheet）

## 功能强大
1.解决导出大量数据造成的内存溢出问题（支持分页查询数据库、采用poi官方推荐api（SXSSFWorkbook）, 实现指定行数刷新到磁盘）

2.解决导入大量数据造成的内存溢出问题（支持分页插入数据库、采用poi官方推荐api（XSSF and SAX），采用SAX模式一行行读取到内存当中去)

3.解决含有占位符的空假行造成的读空值问题

4.解决Long类型或者BigDecimal的精度不准问题

## 组件特色
1.导入可以自定义解析成功或失败的处理逻辑

2.导出支持分页查询、全量查询, 自定义每条数据的处理逻辑

3.内置缓存, 3万条11列数据, 排除查询数据所用时间, 第一次导出2.2s左右、第二次导出在1.4s左右；第一次导入3.5s左右、第二次导入2.5s左右

4.注解操作, 轻量且便捷

5.内置常用正则表达式类RegexConst（身份证号、手机号、金额、邮件）

6.可配置是否适配单元格宽度, 默认开启(单元格内容超过20个汉字不再增加宽度, 3万条11列数据, 耗时50ms左右, 用时与数据量成正比)

7.假如出现异常,Sheet、行、列位置也都一并打印

8.注解中的用户自定义字符串信息以及Excel信息已全部trim,不用担心存在前后空格的风险

9.Excel样式简洁、大方、美观

10.导出的单条数据假如全部属性都为null或0或0.0或0.00或空字符串者null字符串,自动忽略,此特性也可让用户自定义忽略规则

11.除了直接返回OutputStream的方法以外的导出方法, 正常或异常情况都会自动关闭OutputStrem、Workbook流

## 组件需知
### 导入&导出
1.导入和导出只支持尾缀为xlsx的Excel文件

2.标注注解的属性顺序即Excel列的排列顺序

3.时间转化格式(dateFormat)默认为“yyyy-MM-dd HH:mm:ss“

### 导入
1.当导入Excel, 读取到空行, 则停止读取当前Sheet的后面数据行

2.导入Excel文件, 单元格格式使用文本或者常规, 防止出现不可预测异常

3.导入字段类型支持:Date、Short（short）、Integer（int）、Double（double）、Long（long）、Float（float）、BigDecimal、String类型

4.导入BigDecimal字段精度默认为2, roundingMode默认为BigDecimal.ROUND_HALF_EVEN, scale设置为-1则不进行格式化

5.第一行有效单元格内必须包含内容并且以第一行为依据, 导入Excel文件列数必须等于标注注解的属性数量

6.Date类型字段,Excel与时间转化格式(dateFormat)相比,格式要保持一致（反例：2018/12/31和“yyyy-MM-dd“）并且长度要一致或更长（反例："2018-12-31"和yyyy-MM-dd HH:mm:ss"）,否则SimpleDateFormat将解析失败,报 “Unparseable date:”

### 导出
1.导出BigDecimal字段默认不进行精度格式化

2.分页查询默认从第一页开始, 每页3000条

3.Excel每超过2000条数据, 将内存中的数据刷新到磁盘当中去

4.使用分Sheet导出方法, 每8万行数据分Sheet

5.当使用（exportResponse、exportStream、generateExcelStream）方法时, 当单个Sheet超过100万条则会分Sheet

6.标注属性类型要与数据库类型保持一致

7.如果想提高性能, 并且内存允许、并发导出量不大, 可以根据实际场景适量改变分页条数和磁盘刷新量

## 扩展
1.新建子类继承ExcelBoot类, 使用子类构造器覆盖以下默认参数, 作为通用配置

2.直接调用以下两个构造器, 用于临时修改配置
```java
/**
* HttpServletResponse 通用导出Excel构造器
*/
ExportBuilder(HttpServletResponse response, String fileName, Class excelClass, Integer pageSize, Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth)
/**
* OutputStream 通用导出Excel构造器
*/
ExportBuilder(OutputStream outputStream, String fileName, Class excelClass, Integer pageSize, Integer rowAccessWindowSize, Integer recordCountPerSheet, Boolean openAutoColumWidth)
```
```java
    /**
     * Excel自动刷新到磁盘的数量
     */
    public static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 2000;
    /**
     * 分页条数
     */
    public static final int DEFAULT_PAGE_SIZE = 3000;
    /**
     * 分Sheet条数
     */
    public static final int DEFAULT_RECORD_COUNT_PEER_SHEET = 80000;
    /**
     * 是否开启自动适配宽度
     */
    public static final boolean OPEN_AUTO_COLUM_WIDTH = true;
```
## 版本
当前为2.0版本, 新版本正在开发

## 使用手册
1.引入Maven依赖

2.将需要导出或者导入的实体属性上标注@ExportField或@ImportField注解

3.直接调用导出或导入API即可

### POM.xml 

```xml
<dependency>
	<groupId>io.github.magic-core</groupId>
	<artifactId>excel-boot</artifactId>
	<version>2.0</version>
</dependency>
```
### 导出导入实体对象
```java
/**
* 导出导入实体对象
*/
public class UserEntity {
/**
* Integer类型字段
*/
@ExportField(columnName = "ID", defaultCellValue = "1")
@ImportField(required = true)
    private Integer id;
/**
* String类型字段
*/
@ExportField(columnName = "姓名", defaultCellValue = "张三")
@ImportField(regex = IDCARD_REGEX, regexMessage="身份证校验失败")
    private String name;
/**
* BigDecimal类型字段
*/    
@ExportField(columnName = "收入金额", defaultCellValue = "100", scale = 2, roundingMode=BigDecimal.ROUND_HALF_EVEN)
@ImportField(scale = 2, roundingMode=BigDecimal.ROUND_HALF_EVEN)
    private BigDecimal money;
/**
* Date类型字段
*/
@ExportField(columnName = "创建时间", dateFormat="yyyy-MM-dd", defaultCellValue = "2019-01-01")
@ImportField(dateFormat="yyyy-MM-dd")
    private Date birthDayTime;
}
```
### 导出api-Demo
```java
/**
 * 导出api-Demo
 * 
 * UserEntity是标注注解的类,Excel映射的导出类
 * ParamEntity是数据层查询的参数对象
 * ResultEntity是数据层查询到的List内部元素
 * UserEntity可以和ResultEntity使用同一个对象,即直接在数据层查询的结果对象上标注注解(建议使用两个对象, 实现解耦)
 * 
 * pageQuery方法是用户自己实现, 根据查询条件和当前页数和每页条数进行数据层查询
 * convert方法是用户自己实现, 参数就是您查询出来的list中的每个元素引用, 您可以对对象属性的转换或者对象的转换, 如果不进行转换,直接返回参数对象即可
 */
@Controller
@RequestMapping("/export")
public class TestController {
    /**
     * 浏览器导出Excel
     *
     * @param httpServletResponse
     */
    @RequestMapping("/exportResponse")
    public void exportResponse(HttpServletResponse httpServletResponse) {
        ParamEntity queryQaram = new ParamEntity();
        ExcelBoot.ExportBuilder(httpServletResponse, "Excel文件名", UserEntity.class).exportResponse(queryQaram,
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
                    public UserEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 浏览器多sheet导出Excel
     *
     * @param httpServletResponse
     */
    @RequestMapping("/exportMultiSheetResponse")
    public void exportMultiSheetResponse(HttpServletResponse httpServletResponse) {
        ParamEntity queryQaram = new ParamEntity();
        ExcelBoot.ExportBuilder(httpServletResponse, "Excel文件名", UserEntity.class).exportMultiSheetStream(queryQaram,
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
                    public UserEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 导出Excel到指定路径
     */
    @RequestMapping("/exportStream")
    public void exportStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        ExcelBoot.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", UserEntity.class)
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
                    public UserEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 导出多sheet Excel到指定路径
     */
    @RequestMapping(value = "exportMultiSheetStream")
    public void exportMultiSheetStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        ExcelBoot.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", UserEntity.class)
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
                    public UserEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 生成Excel OutputStream对象
     */
    @RequestMapping(generateStream)
    public void generateExcelStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        OutputStream outputStream = ExcelBoot.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", UserEntity.class)
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
                    public UserEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 生成多Sheet Excel OutputStream对象
     */
    @RequestMapping(generateMultiSheetStream)
    public void generateMultiSheetExcelStream() throws FileNotFoundException {
        ParamEntity queryQaram = new ParamEntity();
        OutputStream outputStream = ExcelBoot.ExportBuilder(new FileOutputStream(new File("C:\\Users\\Excel文件.xlsx")), "Sheet名", UserEntity.class)
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
                    public UserEntity convert(ResultEntity o) {
                        //转换操作
                    }
                });
    }

    /**
     * 导出Excel模板
     */
    @RequestMapping("/exportTemplate")
    public void exportTemplate(HttpServletResponse httpServletResponse) {
        ExcelBoot.ExportBuilder(httpServletResponse, "Excel模板名称", UserEntity.class).exportTemplate();
    }
}
```
### 导入api-Demo
```java
/**
 * 导入api-Demo
 * 
 * UserEntity是标注注解的类, Excel映射的导入类, onProcess的userEntity参数则是Excel每行数据的映射实体
 * ErrorEntity是封装了每行Excel数据常规校验后的错误信息实体, 封装了sheet号、行号、列号、单元格值、所属列名、错误信息
 * 
 * onProcess方法是用户自己实现, 当经过正则或者判空常规校验成功后执行的方法,参数是每行数据映射的实体
 * convert方法是用户自己实现, 当经过正则或者判空常规校验失败后执行的方法
 */
@Controller
@RequestMapping("/import")
public class TestController {
    @RequestMapping("/importExcel")
    public void importExcel() throws IOException {
        ExcelBoot.ImportBuilder(new FileInputStream(new File("C:\\Users\\导入Excel文件.xlsx")),  UserEntity.class)
                .importExcel(new ExcelImportFunction<UserEntity>() {

                    /**
                     * @param sheetIndex 当前执行的Sheet的索引, 从1开始
                     * @param rowIndex 当前执行的行数, 从1开始
                     * @param resultEntity Excel行数据的实体
                     */
                    @Override
                    public void onProcess(int sheetIndex,  int rowIndex,  UserEntity userEntity) {
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
**建议功能投票列表**

例：（功能描述）：（票数）