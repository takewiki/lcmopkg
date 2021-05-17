name <-c('B','A','BC','ACF')
fid <- c(1,1,2,3)
mydata <- data.frame(name,fid,stringsAsFactors = F)
mydata2 <- mydata[order(mydata$fid,mydata$name),]
mydata2
