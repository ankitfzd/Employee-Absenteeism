#Clean the environment nd delete all th variables
rm(list = ls())

#Set working Directory
setwd("C:/Users/Ankit singh/Downloads")


#Sys.setenv(JAVA_HOME='C:\\Program Files\\Java\\jre1.8.0_211')
#library(rJava)
x = c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071",
     "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine')
lapply(x, require, character.only = TRUE)
rm(x)


#Reading the data
library(xlsx)
df=read.xlsx("Absenteeism_at_work_Project.xls",sheetIndex=1)

#structure of object
str(df)

#-------------------------------------Exploratory Data Analysis---------------------------

#Looking at the data, we noticed three rows have month of absence as 0. As month can't be 0, we have removed those rows. 
#Another reason for removing this row is the trend between Disciplinary failure and Reason of Absence. Disciplinary failure is always 1 wherever reaon of absence is 0 expcet these three rows
df=df[1:737,]

# Dividing the columns into categorical and continuous

#Categorical variabes
cat=c("ID","Reason.for.absence","Month.of.absence","Day.of.the.week","Seasons", "Social.drinker", "Social.smoker", "Education","Disciplinary.failure","Son", "Pet")        
#
#Continuous variables
cont=c("Transportation.expense","Distance.from.Residence.to.Work","Service.time","Age","Work.load.Average.day.","Hit.target", "Weight", "Height","Body.mass.index" ,"Absenteeism.time.in.hours")


---------------------------------#Missing value analysis----------------------------

#To Calculate missing values present in each variable
missing_val = data.frame(apply(df,2,function(x){sum(is.na(x))}))
missing_val$Columns = row.names(missing_val)
names(missing_val)[1] =  "Missing_percentage"

#To calculate missing value %
missing_val$Missing_percentage = (missing_val$Missing_percentage/nrow(df)) * 100

# Sorting missing_val in Descending order
missing_val = missing_val[order(-missing_val$Missing_percentage),]
row.names(missing_val) = NULL


# Rearranging columns
missing_val = missing_val[,c(2,1)]
missing_val
#Ploting missing value

library(ggplot2)
ggplot(data = missing_val[1:21,], aes(x=reorder(Columns, -Missing_percentage),y = Missing_percentage))+
geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
ggtitle("Missing data percentage (Train)") + theme_bw()

# I am using different techniques to impute different variable. Many of the features like Height, Weight, Distance from work, transportation expenses are unique for each employee and thus can be imputed based on employee id. 


library(simputation)
df=impute_median(df,Transportation.expense+Distance.from.Residence.to.Work+Service.time+Age+Son+Pet+Weight+Height+Body.mass.index+Education+Social.drinker+Social.smoker~ID)
df$Month.of.absence[is.na(df$Month.of.absence)]=10
df=impute_median(df,Work.load.Average.day.~Month.of.absence)#Workload is same for each month


# As Absenteeism time can't be function of ID, so removing it. 
df=subset(df,select=names(df)!=("ID"))
cat=c("Reason.for.absence","Month.of.absence","Day.of.the.week","Seasons", "Social.drinker", "Social.smoker", "Education","Disciplinary.failure","Son", "Pet")        

#Imputing categorical variables using mode

Mode=function(x){
  uniq=unique(x)
  uniq[which.max(tabulate(match(x,uniq)))]
}
remove(mode)
for(i in cat){
  print(i)
  df[,i][is.na(df[,i])] = Mode(df[,i])
}

# The remaining value is populated through kNN 
df=knnImputation(df,k=3)

#Checking the number of missing values 
sum(is.na(df)) #It is equal to 0


#-------------------------------------OUtlier analysis--------------------------------------

#Draw boxplot to detect outliers for continuous variable
for (i in 1:length(cont)){
print(colnames(df[cont[i]]))
  assign(paste0("gn",i), ggplot(aes_string(y = colnames(df[cont[i]]), x = "Absenteeism.time.in.hours"), data = subset(df))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "grey" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=colnames(df[cont[i]]),x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot of ",colnames(df[cont[i]]))))
}    

gridExtra::grid.arrange(gn1, gn2,gn3, ncol=3)
gridExtra::grid.arrange(gn4,gn5,gn6, ncol=3)
gridExtra::grid.arrange(gn7,gn8, ncol =2)
gridExtra::grid.arrange(gn9,gn10,ncol =2 )

#Replace all outliers with NA and impute the missig values
# Since age and HEight are two features whihc are employee sepcific and can take extreme values. Except these two, removing outliers from others
for(i in cont){
print(i)
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
print(length(val))
df[,i][df[,i] %in% val] = NA
}
sum(is.na(df))

# Imputing using KNN
df = knnImputation(df, k = 3)


#---------------------------Feature Selection-------------------------------

## Correlation Plot 
corrgram(df[,cont], order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")



# As weight and Body mass index are highly correlated and thus will carry same information. Removing Body mass index as Weight is a basic variable
df=subset(df,select=names(df)!=("Body.mass.index"))

#Updating cont variable
cont=c("Transportation.expense","Distance.from.Residence.to.Work","Service.time","Age","Work.load.Average.day.","Hit.target", "Weight", "Height" ,"Absenteeism.time.in.hours")

anova_test = aov(Absenteeism.time.in.hours ~ Day.of.the.week + Education + Social.smoker + Social.drinker + Reason.for.absence + Seasons + Month.of.absence + Disciplinary.failure+Son+Pet, data = df)
summary(anova_test)

df=subset(df,select=-c(Education,Social.smoker,Month.of.absence,Pet))
cat=c("Reason.for.absence","Day.of.the.week","Seasons", "Social.drinker","Disciplinary.failure","Son")        
df1=df
df=dummy.data.frame(df,cat)


write.xlsx(df1,"Absenteeism_at_work_Project.xls", sheetName="Sheet2")



#------------------------------------Feature Scaling--------------------
#MAking histogram ti chekc if varables are uniformy distributed

#Normality check
hist(df$Absenteeism.time.in.hours)
hist(d$Transportation.expense)
hist(df$Distance.from.Residence.to.Work)
hist(df$Service.time)
hist(df$Absenteeism.time.in.hours)

# as the variables are not normlly distrubuted, we need to normalize them instead of stndardization
for(i in cont){
  
  print(colnames(df[i]))
  df[[i]] = (df[[i]] - min(df[i]))/
    (max(df[i] - min(df[[i]])))
}



#-----------------------------Model Development-------------
#Clean the environment

set.seed(123)
# PArtioning the data using simple random sampling

train.index = sample(1:nrow(df), 0.8 * nrow(df))
train = df[train.index,]
test  = df[-train.index,]

#---------------------------Dimensaionality reduction using PCA


#As there are lot of variables present fter converting categoricl to dummies especially for Reson of Absence to 
# USing PCA to reduce dimensionality
#install.packages("prcomp")
#library(prcomp)

#caluclating compponents 
prin = prcomp(train)

#Calculating SD for each principal component
sd = prin$sdev

#compute variance
pvar = sd^2

#Calculating how much proportion of variance is explained
var_e = pvar/sum(pvar)

#cumulative scree plot
plot(cumsum(var_e), xlab = "Principal Component",
     ylab = "Cumulative Proportion of Variance Explained",
     type = "b")

#Including data with principal components
train_pc = data.frame(Absenteeism.time.in.hours = train$Absenteeism.time.in.hours, prin$x)

# The below  plot shows that if 40 components explains almost 99+ % data variance
train_pc =train_pc[,1:40]

#transform test into PCA
test_pc = as.data.frame(predict(prin, newdata = test))

#select the first 50 components of test data too
test_pc=test_pc[,1:40]

summary(prin)


#---------------------------------------Descion Tree Regression-------------------

library(rpart)

#Model development
DT_model = rpart(Absenteeism.time.in.hours ~., data = train_pc, method = "anova")


#Prediction for training and testdata
DT_train= predict(DT_model)
DT_test= predict(DT_model,test_pc)

# Metrics for  training data 
print(postResample(pred = DT_train, obs = train$Absenteeism.time.in.hours))
#RMSE       Rsquared        MAE 
#0.09591796 0.76218239 0.06607514

# Metrics for testing data 
print(postResample(pred = DT_test, obs = test$Absenteeism.time.in.hours))
#RMSE        Rsquared        MAE 
#0.12086954 0.73957300 0.07592182

#--------------------------Random Forest Model
library("randomForest")

#Model development
RF_model = randomForest(Absenteeism.time.in.hours~. , train_pc,  ntree=500)


#Prediction for training and testdata
RF_train = predict(RF_model,train_pc)
RF_test=predict(RF_model,test_pc)

# Metrics for  training data  
print(postResample(pred = RF_train, obs = train$Absenteeism.time.in.hours))
#RMSE          R-squared       MAE
#0.03458086   0.98430420   0.02481082 

## Metrics for testing data 
print(postResample(pred = RF_test, obs =test$Absenteeism.time.in.hours))

#RMSE          R-squared       MAE
#0.09673609 0.88095322     0.07115283


#------------Linear Regression---------------

library(usdm)
#LR
vifcor(df[cont], th=0.9)
# No multicollinarity problem.  Correlation is maximum between age and service time (~0.66) which is below 0.9 

#Model development
LR_model = lm(Absenteeism.time.in.hours ~ ., data = train_pc)


#Prediction for training and testdata
LR_train = predict(LR_model,train_pc)
LR_test = predict(LR_model,test_pc)

# Metrics for  training data 
print(postResample(pred = LR_train, obs = train$Absenteeism.time.in.hours))
#RMSE          Rsquared       MAE 
#0.004479807  0.999481244  0.001745014

## Metrics for testing data 
print(postResample(pred = LR_test, obs =test$Absenteeism.time.in.hours))
   #RMSE       Rsquared       MAE 
#0.004615743 0.999614947   0.001980873




#--------------------------------------------------Gradient Boosting-----------------------

#Model development
GBM_model = gbm(Absenteeism.time.in.hours~., data = train_pc, n.trees = 500, interaction.depth = 2)

#Prediction for training and testdata
GBM_train = predict(GBM_model, train_pc, n.trees = 100)
GBM_test = predict(GBM_model,test_pc, n.trees = 100)

# Metrics for  training data  
print(postResample(pred = GBM_train, obs = train$Absenteeism.time.in.hours))
#RMSE       Rsquared       MAE 
#0.05618817 0.92666536 0.04253619

## Metrics for testing data 
print(postResample(pred = GBM_test, obs = test$Absenteeism.time.in.hours))

#RMSE          Rsquared       MAE 
#0.07755080  0.90837434   0.06015437


#The best R-square and least RMSE is of linear regression and thus we will go ahead wih it


# Question 1
------------------------------------# Important features extraction-----------------------

#install.packages("Boruta")

#####Question 1. -------------------------

library(Boruta)
set.seed(123)
df.boruta <- Boruta(Absenteeism.time.in.hours ~., data = df, doTrace = 2)

final <- TentativeRoughFix(df.boruta)
print(final)
final_col=getSelectedAttributes(final)
df2=df[,final_col]
df2$Absenteeism.time.in.hours =df$Absenteeism.time.in.hours
df=df2
#Plotting boruta
plot(final, xlab = "", xaxt = "n")
lz<-lapply(1:ncol(final$ImpHistory),function(i)
  final$ImpHistory[is.finite(final$ImpHistory[,i]),i])
names(lz) <- colnames(final$ImpHistory)
Labels <- sort(sapply(lz,median))
axis(side = 1,las=2,labels = names(Labels),
     at = 1:ncol(final$ImpHistory), cex.axis = 0.52)

#Question 2--------------------------------------------------------------


  df1$Loss= (df1$Work.load.Average.day.*df1$Absenteeism.time.in.hours)/8

#Average_los/month= 

sum(df1$Loss)/12  


final_col
