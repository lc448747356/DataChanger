# DataChanger
change excel data to a format that can be used in unity parsing

a.在UNITY中生成txt数据文件
1.第一行为字段名称，第二行可以添加备注，第三行开始为正常数据
2.列添加备注可在该列的第一行的字段名称处填写remark则该行不会生成相应数据
3.table名称请勿使用默认名称 如 Sheet1 Sheet2 Sheet3,且应该与项目中的解析的模型名称相对应
4.支持解析类型int,string,float,float[],int[]    (其中数组类型用;隔开)
5.左侧ID对应游戏内解析的Dictionary ID,不同的table可以使用相同的ID，相同的table不可以使用同一ID

b.在UNITY中对txt文件进行解析
1.先在DataManager中添加相应词典，名称为   _类型名称 如 test，则名称为_test
2.解析出的数据类型由模型确定 暂时支持int,string,float,float[],int[]
