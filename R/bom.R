library(readxl)

#' Title bom展开函数
#'
#' @param file_name  文件名
#' @param sheet_data data页签
#' @param sheet_BOM BOM页签
#'
#' @return  返回值
#' @export
#'
#' @examples
#' bom_breakdown()
bom_breakdown <- function(file_name="data-raw/需求说明文档LD.xlsx",sheet_data="2023新模板表",sheet_BOM="2023BOM表") {


  data <- readxl::read_excel(file_name,sheet = sheet_data)
  data2 = data[data$物料描述 == "子箱" , ]
  dataB = data[data$物料描述 != "子箱" , ]
  df1 = data2$采购凭证号
  data3 <- readxl::read_excel(file_name,sheet =sheet_BOM)

  result=lapply(df1, function(item){
    data4 = data3[data3$采购订单号 == item, ]
    #print(data4)   此处可正常遍历data4，print写在循环外不显示遍历结果
    names(data4)

    col_selected = c('BOM组件','物料描述','物料长文本','图号','G番','L番','变量','版本号','互换性','组件数量','BOM项目文本','有无摘要','VDA')
    data5 = data4[ ,col_selected]

    data2_bom = data2[data2$采购凭证号 ==item, ]
    #注：变量item不可加引号
    data5$主订单号= data2_bom$主订单号

    data5$采购凭证号 = data2_bom$采购凭证号

    data5$订购日期 = data2_bom$订购日期

    data5$供应商编码 = data2_bom$供应商编码
    data5$供应商名称 = data2_bom$供应商名称

    data5$采购组 = data2_bom$采购组

    data5$新品 = data2_bom$新品

    data5$采购分类 = data2_bom$采购分类

    data5$交货期 = data2_bom$交货期

    data5$单价 = data2_bom$单价

    data5$单价类型 = data2_bom$单价类型

    data5$币种 = data2_bom$币种

    data5$单位 = data2_bom$单位

    data5$工事番号 = data2_bom$工事番号
    data5$客户订单号 = data2_bom$客户订单号

    data5$项目名称 = data2_bom$项目名称

    data5$排产交货期 = data2_bom$排产交货期

    data5$工事番号 = data2_bom$工事番号
    data5$客户订单号= data2_bom$客户订单号

    data5$项目名称 = data2_bom$项目名称

    data5$工事番号 = data2_bom$工事番号
    data5$排产交货期 = data2_bom$排产交货期
    data5$成组信息 = data2_bom$成组信息

    data5$仓号 = data2_bom$仓号

    data5$箱号= data2_bom$箱号

    data5$采购凭证类型 = data2_bom$采购凭证类型
    data5$采购组织= data2_bom$采购组织

    data5$工厂= data2_bom$工厂

    data5$交货场所= data2_bom$交货场所

    data5$库存地点= data2_bom$库存地点

    data5$熔断标识= data2_bom$熔断标识

    data5$采购凭证中的项目类别= data2_bom$采购凭证中的项目类别

    data5$订单状态= data2_bom$订单状态
    data5$采购申请单号= data2_bom$采购申请单号
    data5$物料号=data5$BOM组件
    data5$变数=data5$变量
    data5$订购数量=data5$组件数量
    data5$物料组=data5$BOM项目文本
    data5 = data5[ ,c('主订单号','采购凭证号','订购日期','物料号','物料描述','供应商编码','供应商名称','采购组','物料长文本','图号','G番','L番','变数','版本号','互换性','新品','采购分类','订购数量','交货期','单价','单价类型','币种','单位','物料组','工事番号','客户订单号','项目名称','排产交货期','有无摘要','VDA','成组信息','仓号','箱号','采购凭证类型','采购组织','工厂','交货场所','库存地点','熔断标识','采购凭证中的项目类别','订单状态','采购申请单号')]
    return(data5)


  })
  dataA=do.call('rbind',result)
  data=rbind(dataA,dataB)
  data$类别 = data$采购分类
  data$生产旬 <-''
  data$主订单编号 <- data$主订单号
  data$物料号 <- data$物料号
  data$采购单价 <-data$单价
  data$单价区分 <- data$单价类型
  data$子订单编号 <- data$采购凭证号
  data$品目 <- data$物料组
  #针对数据进行预处理
  data$图号 <- tsdo::na_replace(data$图号,'')
  data$G番 <- tsdo::na_replace(data$G番,'')
  data$L番 <- tsdo::na_replace(data$L番,'')
  data$图号 <- paste(data$图号,data$G番,data$L番,sep = ' ')
  data$图号版本号 <- tsdo::na_replace(data$版本号,'')
  data$资材编号 <-''
  data$品名 <- data$物料描述
  data$库区 <-''
  data$`材质规格/尺寸` <-''
  data$VDA信息 <- data$VDA
  data$供应商 <- data$供应商编码
  data$采购员 <-''
  data$客户号 <- data$客户订单号
  data$客户名称 <- data$项目名称
  data$成组编号 <- data$交货场所
  data$子订单状态 <- data$订单状态
  data$APD标志位 <-''

  col_selected <- c('类别',
                    '生产旬',
                    '主订单编号',
                    '子订单编号',
                    '订购日期',
                    '品目',
                    '仓号',
                    '物料号',
                    '图号',
                    '图号版本号',
                    '互换性',
                    '新品',
                    '资材编号',
                    '变数',
                    '品名',
                    '库区',
                    '材质规格/尺寸',
                    'VDA信息',
                    '供应商',
                    '供应商名称',
                    '工事番号',
                    '采购员',
                    '交货期',
                    '订购数量',
                    '采购单价',
                    '单价区分',
                    '客户号',
                    '客户名称',
                    '有无摘要',
                    '成组编号',
                    '子订单状态',
                    'APD标志位'
  )
  data = data[ ,col_selected]
  return(data)

}


#
# openxlsx::write.xlsx(result,'result.xlsx')
# print(result)
