
#调用程序???
library(stats)
library(plyr)
library(survival)
library(xlsx)

#载入数据表格

Base <- read.xlsx(data,sheetIndex = 1,header = T,encoding = 'UTF-8')

# 2频率统计
# 2.1 频数
Freq <- lapply(Base[,1:ncol(Base)],table)
# 2.2 频率统计
Prop <- lapply(Freq[1:ncol(Base)],prop.table)
Char <-NULL
for(i in 1:ncol(Base)){Character <- c(names(Freq[i]),names(Freq[[i]]))
Noc <- c(NA,paste0(Freq[[i]],'(',Prop[[i]]*100,')'))
Characteristics <- data.frame('Characteristics' = Character,'Number of Cases(%)' = Noc)
Char <- rbind(Char,Characteristics)
}


# 2.1 输出
write.xlsx(Char,'freq.xlsx',col.names = T,row.names = F,showNA = T)

# 3风险模型
# 3.1转换数据类型、给年龄赋予梯度大小

Base[,c(1:ncol(Base))] <- lapply(Base[,c(1:ncol(Base))],as.factor)
Base$Age.m. <- factor(Base$Age.m.,ordered = T,levels = c('L3','M4','M1'))

# 3.2 单因素回归（批量计算???
Ba <- Base[,1]
UniLo <- function(x){
  FML <- as.formula(paste0('Ba~',x))
  GLo <- glm(FML,family=binomial(logit),data = Base)
  GSum <- summary(GLo)
  HR <- round(exp(coef(GLo))[2],2)
  PValue <- round(round(GSum$coefficients[,4],3)[2],7)
  CI <- paste0(round(round(exp(confint(GLo))[,1],2)[2],2),'-',round(round(exp(confint(GLo))[,2],2)[2],2))
  UniLo <- data.frame('Characteristics' = x,
                      'Hazard Ratio' = HR,
                      'CI95' = CI,
                      'P Value' = PValue)
  return(UniLo)
}

#看所需要的变量在第几列
VarNames <- colnames(Base)[c(2:ncol(Base))]
#计算
UniVar <- lapply(VarNames,UniLo)
#做成数据???
Univar <- ldply(UniVar,data.frame)

#筛选p???(根据自己需要修改后面那个数???)
no <- Univar$Characteristics[Univar$P.Value < 0.2]

#5.3多因素回归分???
fml <- as.formula(paste0('Ba~',paste0(Univar$Characteristics[Univar$P.Value < 0.2],collapse = '+')))
MultiLo <- glm(fml,family=binomial(logit),data = Base)
MultiSum <- summary(MultiLo)

MultiName <- as.character(Univar$Characteristics[Univar$P.Value < 0.2])
MHR <- round(exp(coef(MultiLo))[2:(length(no)+1)],2)
MPV <- round(round(MultiSum$coefficients[,4],3)[2:(length(no)+1)],3)
MCI <- paste0(round(round(exp(confint(MultiLo))[,1],2)[2:(length(no)+1)],2),'-',round(round(exp(confint(MultiLo))[,2],2)[2:(length(no)+1)],2))

MultiLo <- data.frame('Characteristics' = MultiName,
                      'Hazard Ratio' = MHR,
                      'CI95' = MCI,
                      'P Value' = MPV)

# 5.4输出
Univarname <- paste0(data,'-','Univar.xlsx',seq='')
MultiLoname <- paste0(data,'-','MultiLo.xlsx',seq='')
write.xlsx(Univar,Univarname,col.names = T,row.names = F,showNA = T)
write.xlsx(MultiLo,MultiLoname,col.names = T,row.names = F,showNA = T)

