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

## Getting started
- 将target目录生成的hugb-tools-x.x.x.jar包部署到%PS_HOME%/class目录下(不建议)或者配置psappsrv.cfg和psprcs.cfg文件中Add to CLASSPATH=指定的目录中;


- 重启AppServer和ProcessServer

## Excel report
``` java
/*==========================================================
  任务编号: 报表输出
  说   明: Excel类
  ---------------------------------------------------------
  日期             	作者 	        	说明
  2019-06-24 	        hugb    	        创建
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