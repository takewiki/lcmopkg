library(readxl)
data <- read_excel("data-raw/新三菱订单模板2023转2022.xlsx",
                   sheet = "2023新模板表")
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

View(data)

