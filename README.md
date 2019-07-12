# hugb-peoplesoft-tools
peoplesoft常用工具类汇总

## Development
项目依赖[maven](http://maven.apache.org/install.html)环境，请先执行mvn -v命令检查当前mvn环境配置 ; 

```bash
# clone the project
git clone git@github.com:bughuhu/hugb-tools.git

# cd hugb-tools and build package
mvn clean package
```

## 入门
- 将target目录生成的hugb-tools-x.x.x.jar包部署到%PS_HOME%/class目录下(不建议)或者配置psappsrv.cfg和psprcs.cfg文件中Add to CLASSPATH=指定的目录中;


- 重启AppServer和ProcessServer

## Excel报表

- PeopleSoft Application Design中新建ExcelWrite类;该类封装了excel输出的jar包调用;

``` PeopleCode
/*==========================================================
  任务编号: 报表输出
  说   明: Excel类
  ---------------------------------------------------------
  日期             	作者 	        	说明
  2019-06-24 	        softworm    	        创建
===========================================================*/
class ExcelWrite
   /*构造函数*/
   method ExcelWrite(&reportName As string);
   
   /*设置报表参数*/
   method buildReportParameter();
   /*设置报表表头*/
   method buildReportHeader();
   /*设置报表数据*/
   method buildReportData();
   /*输出报表*/
   method publish();
   
   /*报表参数*/
   property array of array of any parameter get set;
   /*报表头*/
   property array of array of any header get set;
   /*报表数据*/
   property array of array of any data get set;
   
private
   
   /*获取制表人姓名*/
   method getOperatorName() Returns string;
   
   instance string &_reportName;
   instance JavaObject &reportDefn;
   instance JavaObject &reportRow;
   
   instance array of array of any &_parameter;
   instance array of array of any &_header;
   instance array of array of any &_data;
end-class;

/*构造函数*/
method ExcelWrite
   /+ &reportName as String +/
   /*报表名*/
   &_reportName = &reportName;
   /*报表实例*/
   &reportDefn = CreateJavaObject("cn.hugb.excel.ReportDefinition");
   
   &_parameter = CreateArray(CreateArrayAny());
   &_header = CreateArray(CreateArrayAny());
   &_data = CreateArray(CreateArrayAny());
   
end-method;

/*设置报表参数*/
method buildReportParameter
   /*报表名*/
   &reportDefn.setReportName(%This._reportName);
   /*制表人姓名*/
   &reportDefn.setOperatorName(%This.getOperatorName());
   
   Local integer &i;
   For &i = 1 To &_parameter.Len
      /*追加报表参数*/
      &reportDefn.addParameter(&_parameter [&i][1], &_parameter [&i][2]);
   End-For;
end-method;

/*设置报表表头*/
method buildReportHeader
   Local integer &i;
   For &i = 1 To &_header.Len
      /*追加报表头*/
      &reportDefn.addHeader(&_header [&i][1]);
   End-For;
   
end-method;

/*设置报表数据*/
method buildReportData
   
   Local string &fieldType;
   Local integer &i, &j;
   
   /*判断表头是否设置*/
   If &_header.Len <= 0 Then
      throw CreateException(0, 0, "请设置报表头参数[header].");
   End-If;
   
   For &i = 1 To &_data.Len
      
      /*行实例*/
      &reportRow = CreateJavaObject("cn.hugb.excel.ReportRow");
      
      For &j = 1 To &_header.Len
         
         /*根据报表头参数2判断类型*/
         try
            &fieldType = &_header [&j][2];
         catch Exception &ex
            /*异常一律按String处理*/
            &fieldType = "String";
         end-try;
         
         Evaluate &fieldType
            /*Number*/
         When "Number"
            &reportRow.addNumber(String(&_data [&i][&j]));
            Break;
            /*Date*/
         When "Date"
            &reportRow.addDate(String(&_data [&i][&j]));
            Break;
            /*其他类型，全部按String处理*/
         When-Other
            &reportRow.addString(String(&_data [&i][&j]));
         End-Evaluate;
      End-For;
      
      /*追加到报表*/
      &reportDefn.addRow(&reportRow);
   End-For;
end-method;

/*输出报表*/
method publish

   /*设置报表参数*/
   %This.buildReportParameter();
   
   /*设置报表表头*/
   %This.buildReportHeader();
   
   /*设置报表数据*/
   %This.buildReportData();
   
   /*报表路径*/
   Local string &filePath = %FilePath | %This._reportName | "_" | UuidGen() | ".xls";
   MessageBox(0, "", 0, 0, &filePath);
   &reportDefn.buildExcel(&filePath);
end-method;


/*获取制表人姓名*/
method getOperatorName
   /+ Returns String +/
   Local string &operatorName;
   SQLExec("SELECT O.OPRDEFNDESC FROM PSOPRDEFN O WHERE O.OPRID=:1 AND O.OPRDEFNDESC<>' '", %OperatorId, &operatorName);
   If All(&operatorName) Then
      Return &operatorName;
   End-If;
   Return %OperatorId;
end-method;

set parameter
   /+ &NewValue as Array2 of Any +/
   &_parameter = &NewValue;
end-set;

get parameter
   /+ Returns Array2 of Any +/
   Return &_parameter;
end-get;

set header
   /+ &NewValue as Array2 of Any +/
   &_header = &NewValue;
end-set;

get header
   /+ Returns Array2 of Any +/
   Return &_header;
end-get;

set data
   /+ &NewValue as Array2 of Any +/
   &_data = &NewValue;
end-set;

get data
   /+ Returns Array2 of Any +/
   Return &_data;
end-get;
```

- 使用样例

```java
/*==========================================================
  任务编号: XXX
  说    明: XX结果合并报表数据源
  ---------------------------------------------------------
  日期             	作者 	        	说明
  2019-06-24 	        softworm     	        创建
===========================================================*/
import XX_XX_XX:ExcelWrite;

class ResultReportDS
   /*构造函数*/
   method ResultReportDS();
   
   /*初始化报表数据源*/
   method initRptDataSource(&paramRec As Record) Returns array of array of any;
   
   /*获取数据*/
   method getRptDataSource(&paramRec As Record) Returns array of array of any;
   
   /*获取数据源SQL*/
   method getRptDataSourceSQL(&paramRec As Record) Returns string;
   
   /*获取报表参数*/
   method getRptParameter(&paramRec As Record) Returns array of array of any;
   
   /*获取报表表头*/
   method getRptHeader(&paramRec As Record) Returns array of array of any;
   
   /*报表发布*/
   method publishReport(&paramArray As array of array of any, &headerArray As array of array of any, &dsArray As array of array of any, &paramRec As Record);
   
private
   instance QH_ABS_CAL:ExcelWrite &excelWrite;
end-class;

/*构造函数*/
method ResultReportDS
   &excelWrite = create QH_ABS_CAL:ExcelWrite("QH_ABS_001");
end-method;

method initRptDataSource
   /+ &paramRec as Record +/
   /+ Returns Array2 of Any +/
   Local array of array of any &dsArray = %This.getRptDataSource(&paramRec);
   Return &dsArray;
end-method;

method getRptDataSource
   /+ &paramRec as Record +/
   /+ Returns Array2 of Any +/
   Local array of array of any &dsArray = CreateArrayRept(CreateArrayAny(), 0);
   
   Local array of any &temp = CreateArrayAny();
   Local array of any &dsTempArray;
   Local SQL &sql = CreateSQL(%This.getRptDataSourceSQL(&paramRec));
   Local number &c = 0;
   Local number &index = 0;
   While &sql.Fetch(&temp)
      &dsTempArray = CreateArrayAny();
      &c = &c + 1;
      &dsTempArray.Push(String(&c));
      &dsTempArray.Push(&temp);
      &dsArray.Push(&dsTempArray);
   End-While;
   &sql.Close();
 
   MessageBox(0, "", 0, 0, "len:" | &dsArray.Len);
   Return &dsArray;
end-method;



/*获取数据源SQL*/
method getRptDataSourceSQL
   /+ &paramRec as Record +/
   /+ Returns String +/
   
   Local integer &i;
   Local string &str, &whereBuild;
   
   &str = &str | "                     select t.emplid,                                                        ";
   &str = &str | "                            t.empl_rcd,                                                      ";
   &str = &str | "                            t.gp_paygroup,                                                   ";
   &str = &str | "                            t.empl_class,                                                    ";
   &str = &str | "                            t.location,                                                      ";
   &str = &str | "                            t.cal_prd_id,                                                    ";
   &str = &str | "                            t.schedule_id,                                                   ";
   &str = &str | "                            t.hr_status,                                                     ";
   &str = &str | "                            t.qh_wa_pay_011 /*饭补*/                                         ";
   &str = &str | "                       from ps_qh_abs_rslt_tbl t where 1=1                                   ";
   &str = &str | "                       and t.cal_prd_id=" | Quote(&paramRec.CAL_PRD_ID.Value);
   

   /*安全性控制*/
   &whereBuild = &whereBuild | " AND EXISTS (SELECT 1 FROM PS_DEPT_TBL_ACCESS ACL WHERE ACL.SETID='XXX' AND ACL.DEPTID=T.DEPTID AND ACL.OPRID=" | Quote(%OperatorId) | ")  ";
   
   /*拼接条件*/
   &str = &str | &whereBuild;
   
   Return &str;
end-method;



/*获取报表参数*/
method getRptParameter
   /+ &paramRec as Record +/
   /+ Returns Array2 of Any +/
   Local array of array of any &dsArray = CreateArrayRept(CreateArrayAny(), 0);
   Return &dsArray;
end-method;

/*获取报表表头*/
method getRptHeader
   /+ &paramRec as Record +/
   /+ Returns Array2 of Any +/
   Local integer &i;
   Local array of array of any &dsArray = CreateArrayRept(CreateArrayAny(), 0);
   Local array of any &dsTempArray;
   
   &dsArray.Push("序号");
   &dsArray.Push("员工编号");
   &dsArray.Push("员工记录号");
   &dsArray.Push("薪资组");
   &dsArray.Push("员工类别");
   &dsArray.Push("办公地点");
   &dsArray.Push("期间ID");
   &dsArray.Push("班次计划");
   &dsArray.Push("HR状态");
   
   /*饭补*/
   &dsTempArray = CreateArrayAny();
   &dsTempArray.Push("饭补");
   &dsTempArray.Push("Number");
   &dsArray.Push(&dsTempArray);
   
   Return &dsArray;
end-method;


/*报表发布*/
method publishReport
   /+ &paramArray as Array2 of Any, +/
   /+ &headerArray as Array2 of Any, +/
   /+ &dsArray as Array2 of Any, +/
   /+ &paramRec as Record +/
   
   /*设置报表参数*/
   &excelWrite.parameter = &paramArray;
   /*设置报表头*/
   &excelWrite.header = &headerArray;
   /*数据*/
   &excelWrite.data = &dsArray;
   /*输出报表*/
   &excelWrite.publish();
end-method;

```

- 生成报表，%This.txnRec为Application Engine的AET状态表

```java
   Local array of array of any &dsArray;
   Local array of array of any &paramArray;
   Local array of array of any &headerArray;
   
   /*报表工具类*/
   Local XX_XX_XX:DataSource:ResultReportDS &dsService = create XX_XX_XX:DataSource:ResultReportDS();
   
   /*报表参数*/
   &paramArray = &dsService.getRptParameter(%This.txnRec);
   
   /*报表头*/
   &headerArray = &dsService.getRptHeader(%This.txnRec);
   
   /*数据*/
   &dsArray = &dsService.initRptDataSource(%This.txnRec);
   
   /*发布报表*/
   &dsService.publishReport(&paramArray, &headerArray, &dsArray, %This.txnRec);
```