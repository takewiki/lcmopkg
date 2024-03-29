#' 针对生产订单按工事番号进行合并
#'
#' @param file_name 文件名
#' @param key_word 关键词
#'
#' @return 返回值
#' @export
#'
#' @examples
#' mo_combine()
mo_combine <- function(file_name ="data-raw/原始数据.xlsx",key_word='轿顶站') {

  data_mo <- readxl::read_excel(file_name)
  #选择数据行
  data_mo$flag <- stringr::str_detect(data_mo$`品名`,key_word)
  col_name_selected <- c('生产旬',
                         '主订单编号',
                         '订购日期',
                         '品目',
                         '图号',
                         '品名',
                         '工事番号',
                         '交货期',
                         '订购数量'
  )
  #选择列
  data_mo <- data_mo[data_mo$flag ==TRUE ,col_name_selected]
  ncount <- nrow(data_mo)
  if (ncount >0){
    #排序
    data_mo <-data_mo[order(data_mo$工事番号), ]
    #分组处理
    data_split <- split(data_mo,data_mo$工事番号)
    #按工事番号进行处理
    data_res <- lapply(data_split, function(data){
      prd_month <- unique(data$`生产旬`)[1]
      mo_no <- unique(data$`主订单编号`)[1]
      pur_date <- unique(data$`订购日期`)[1]
      prd_number <- paste(unique(data$`品目`),collapse = '\n')
      chart_no_cell <-unique(data$`图号`)
      chart_len <- nchar(chart_no_cell)
      chart_no_data <- data.frame(chart_no_cell,chart_len,stringsAsFactors = F)
      #按长度排序，然后按顺序
      chart_no_data <- chart_no_data[order(chart_no_data$chart_len,chart_no_data$chart_no_cell), ]
      chart_no_raw <- chart_no_data$chart_no_cell



      chart_no <- paste(chart_no_raw,collapse = '\n')

      prd_name <- paste(unique(data$`品名`),collapse = '\n')
      mo_no2 <-  unique(data$`工事番号`)[1]
      issue_date <- unique(data$`交货期`)[1]
      prd_qty <- 1
      data_cell <- data.frame(prd_month,mo_no,pur_date,prd_number,chart_no,prd_name,mo_no2,issue_date,prd_qty,stringsAsFactors = F)
      names(data_cell) <- col_name_selected
      return(data_cell)
    })
    #针对数据进行标准化处理
    res <- do.call('rbind',data_res)
    rownames(res) <- NULL
  }else{
    res <- data.frame(info='无结果',stringsAsFactors = F)
  }


  return(res)
 # openxlsx::write.xlsx(res,'lcmo.xlsx')

}
