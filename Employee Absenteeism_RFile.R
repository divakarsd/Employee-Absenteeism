#Clear R environment of all the variables
rm(list = ls())

#Set working directory where my data is stored in the system
setwd("C:/Users/swat/Desktop/Working/Employee-Absenteeism-2")

#check whether path is set or not 
getwd()

#Install required libraries
required_lib = c("readxl","ggplot2","gridExtra","corrplot","tree","caret","randomForest","usdm")
install.packages(required_lib)

#Load required libraries
lapply(required_lib,require,character.only = TRUE)

#load data 
df = read_excel("Absenteeism_at_work_Project.xls")
View(df)
class(df)
df = as.data.frame(df)
class(df)

#Check structure of the data 
str(df)

#Check summary of the data
summary(df)

#column names
colnames(df)

## Let us separate categorical data and numerical data 
## on analysing the data and as mentioned in the problem statement we observe following feature set is categorical : reason for absence ,month of absence,day of the week
## season ,disciplinary failure ,education,social drinker,social smoker
## We are explicilty defining categorical variables as the data types for all the features are numeric

numerical_set = c("ID","Transportation expense","Distance from Residence to Work","Service time","Age","Work load Average/day","Hit target","Son","Pet","Height","Weight","Body mass index","Absenteeism time in hours")
categorical_set = c("Reason for absence","Month of absence","Day of the week","Seasons","Disciplinary failure","Education","Social drinker","Social smoker")

#Convert datatype of categorical data into factor levels 
for (i in categorical_set){
  df[,i] = as.factor(df[,i])
}

############## Missing Value Analysis #######################
## we evaluate the percentage of missing value in each column .
## On viewing the dataset ,we see certain entries have 0 value which is not an acceptable entry logically .Hence we replace them as NA in those entries

for(i in c("Reason for absence","Month of absence","Day of the week","Seasons","Education","ID","Age","Weight","Height","Body mass index")){
  df[i][(df[i] == 0)] = NA
}

## Calculate total sum of missing values column wise
## Calculate total sum of missing values column wise

missing_val_column_wise = data.frame(apply(df,2,function(x){sum(is.na(x))}))
names(missing_val_column_wise)[1] = "NA_Sum"
missing_val_column_wise$NA_Percent = (missing_val_column_wise$NA_Sum/nrow(df))*100

##If any column has more than 30% of missing value ,we will drop that column and update numerical and categorical set accordingly

for(i in 1:length(df)){
  if(missing_val_column_wise$NA_Percent[i]>=30){
    df = df[,-i]
    if(any(numerical_set %in% colnames(df[i]))){
      numerical_set = numerical_set[-(grep(colnames(df[i],numerical_set)))]
    }
    else if(any(categorical_set %in% colnames(df[i]))){
      categorical_set = categorical_set[-(grep(colnames(df[i],categorical_set)))]
    }
  }
}

## none of the columns has more than 30% missing value.Thus all the columns are retained for now.

## We will check which imputation method suits best in our data set model
## To check accuracy of our method ,let us impute NA values in row 10 and check with our imputated values with actual values
## Let us impute missing values 
## Mean method for numerical data 
## Mode method for categorical data

#mode
getmode = function(x){
  unique_x = unique(x)
  mode_val = which.max(tabulate(match(x,unique(x))))
}


#Mean 
impute_mean_mode = function(data_set){
  for(i in colnames(data_set)){
    if(sum(is.na(data_set[,i]))!=0){
      if(is.numeric(data_set[,i])){
        data_set[is.na(data_set[,i]),i] = round(mean(data_set[,i],na.rm = TRUE))
      }
      else if(is.factor(data_set[,i])){
        data_set[is.na(data_set[,i]),i] = getmode(data_set[,i])
      }
    }
  }
  return(data_set)
}


#Median 
impute_median_mode = function(data_set){
  for(i in colnames(data_set)){
    if(sum(is.na(data_set[,i]))!=0){
      if(is.numeric(data_set[,i])){
        data_set[is.na(data_set[,i]),i] = median(data_set[,i],na.rm = TRUE)
      }
      else if(is.factor(data_set[,i])){
        data_set[is.na(data_set[,i]),i] = getmode(data_set[,i])
      }
    }
  }
  return(data_set)
}


#KNN
impute_knn = function(data_set){
  install.packages("DMwR")
  library(DMwR)
  if(sum(is.na(data_set))!=0){
    print("Imputing missing value in the data")
    for(i in categorical_set){
      data_set[,i] = as.numeric(data_set[,i])
    }
    data_set = knnImputation(data_set,k=5)
    data_set = round(data_set)
  }
  return(data_set)
}


#Select best method
df = impute_median_mode(df)

################################ Data visualisation #########################################

##numeric data 
hist(df$ID,col = "blue",main = "ID frequency",xlab = "ID",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$"Transportation expense",col = "blue",main = "Transportation expense frequency",xlab = "Transportation expense",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$"Distance from Residence to Work",col = "blue",main = "Dist from Residence to Work frequency",xlab = "Distance from Residence to Work",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$"Service time",col = "blue",main = "Service time frequency",xlab = "Service time",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$Age,col = "blue",main = "Age frequency",xlab = "Age",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$"Hit target",col = "blue",main = "Hit target frequency",xlab = "Hit target",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$Son,col = "blue",main = "Son frequency",xlab = "Son",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$Pet,col = "blue",main = "Pet frequency",xlab = "Pet",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$Height,col = "blue",main = "Height frequency",xlab = "Height",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$Weight,col = "blue",main = "Weight frequency",xlab = "Weight",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$"Body mass index",col = "blue",main = "Body Mass Index frequency",xlab = "Body Mass Index",breaks = 30,cex.lab = 1,cex.axis = 1)
hist(df$"Absenteeism time in hours",col = "blue",main = "Absent hours frequency",xlab = "Absent hours",breaks = 30,cex.lab = 1,cex.axis = 1)
dev.off()


##Category data
barplot(table(df$"Reason for absence"),col = "blue",main = "Reason for absence frequency",xlab = "Reason for absence",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(df$"Month of absence"),col = "blue",main = "Month of absence frequency",xlab = "Month of absence expense",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(df$"Day of the week"),col = "blue",main = "Day of the week frequency",xlab = "Day of the week",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(df$Seasons),col = "blue",main = "Seasons frequency",xlab = "Seasons time",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(df$"Disciplinary failure"),col = "blue",main = "Disciplinary failure frequency",xlab = "Disciplinary failure",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(df$Education),col = "blue",main = "Education frequency",xlab = "Education",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(df$"Social drinker"),col = "blue",main = "Social drinker frequency",xlab = "Social drinker",ylab = "Frequency",cex.lab = 1,cex.axis = 1)
barplot(table(df$"Social smoker"),col = "blue",main = "Social smoker frequency",xlab = "Social smoker",ylab = "Frequency",cex.lab = 1,cex.axis = 1)


##Boxplot of categorical data set vs target variable
dev.off()
for(i in 1:length(categorical_set)){
  assign(paste0("gg",i),ggplot(aes_string(y=df$"Absenteeism time in hours",x=df[,categorical_set[i]]),data = subset(df))
         + stat_boxplot(geom = "errorbar",width = 0.3) +
           geom_boxplot(outlier.colour = "red",fill = "blue",outlier.shape = 18,outlier.size = 1) +
           labs(y = "Absenteeism time in hours",x=names(df[categorical_set[i]])) + 
           ggtitle(names(df[categorical_set[i]])))
  
}
gridExtra::grid.arrange(gg1,gg2,nrow = 2,ncol=1)
gridExtra::grid.arrange(gg3,gg4,nrow = 2,ncol = 1)
gridExtra::grid.arrange(gg5,gg6,nrow = 2,ncol = 1)
gridExtra::grid.arrange(gg7,gg8,nrow = 2,ncol = 1)


##Scatter plot of numerical data set vs the target variable
plot(df$ID,df$"Absenteeism time in hours",xlab = "ID",ylab = "Absenteeism time in hours",main = "Absenteeism time vs ID",col="blue")


####Outlier
for(i in numerical_set){
  val = df[,i][df[,i] %in% boxplot.stats(df[,i])$out]
  print(length(val))
  df[,i][df[,i] %in% val] = NA
}


################################################# Outlier Analysis ################################

## Replace outliers in numerical dataset with NAs using boxplot method 
for(i in numerical_set){
  outlier_value = boxplot.stats(df[,i])$out
  print(names(df[i]))
  print(outlier_value)
  df[which(df[,i] %in% outlier_value),i] = NA
}

## impute missing values with the best method.
df = impute_median_mode(df)

##############################Feature Selection ####################################################

## Correlation analysis 
dev.off()
pairs(df[numerical_set[c(1:6,13)]])
pairs(df[numerical_set[7:13]])

correlation_matrix = cor(df[,numerical_set])
dev.off()
corrplot(correlation_matrix,method = "number",type = 'lower')

## we see only weight and BMI has higher correlation .Thus we will drop one of the variables i.e weight
### using VIF for multicollinearity check 


##Checking multicollinearity
##If VIF>5 thus we can remove these features
vif(df[,numerical_set])

### Using ANOVA test for categorical data 
for(i in categorical_set){
  print(i)
  aov_summary = summary(aov(df$"Absenteeism time in hours"~df[,i],data = df))
  print(aov_summary)
}

##Dimension reduction
##observing p value of the table ,we conclude that we should drop features whose p value is greater than 0.05
df = subset(df, select = -(which(names(df) %in% c("Weight","Day of the week","Seasons","Education","Social smoker","Social drinker"))))


############################## Model Development##################################################

names(df)[names(df) == "Absenteeism time in hours"] <- "Absenteeism_time_in_hours"
names(df)[names(df) == "Month of absence"] <- "Month_of_absence"
names(df)[names(df) == "Transportation expense"] <- "Transportation_expense"
names(df)[names(df) == "Distance from Residence to Work"] <- "Distance_from_Residence_to_Work"
names(df)[names(df) == "Service time"] <- "Service_time"
names(df)[names(df) == "Hit target"] <- "Hit_target"
names(df)[names(df) == "Work load Average/day"] <- "Work_load_Average_per_day"
names(df)[names(df) == "Body mass index"] <- "Body_mass_index"
names(df)[names(df) == "Reason for absence"] <- "Reason_for_absence"
names(df)[names(df) == "Disciplinary failure"] <- "Disciplinary_failure"

train = df
colnames(train)

#create sampling and divide data into train and test
set.seed(123)
train_index = sample(1:nrow(train), 0.8 * nrow(train))

train1 = train[train_index,]
test1 = train[-train_index,]


############################################Decision Tree#########################

library(rpart)
fit = rpart(Absenteeism_time_in_hours ~. , data = train1, method = "anova", minsplit=5)
predictions_DT = predict(fit, test1[,-15])
RMSE(test1[,15], predictions_DT)
summary(fit)


################################Random Forest######################################

library(randomForest)
RF_model = randomForest(Absenteeism_time_in_hours ~. , train1, importance = TRUE, ntree=100)
RF_Predictions = predict(RF_model, test1[,-15])
RMSE(test1[,15], RF_Predictions)
importance(RF_model, type = 1)


##########################Linear Regression#########################################

lm_model = lm(Absenteeism_time_in_hours ~. , data = train1)
predictions_LR = predict(lm_model, test1[,-15])
RMSE(test1[,15], predictions_LR)
summary(lm_model)


#########################KNN Implementation###########################################

library(class)
KNN_Predictions = knn(train1[,-15], test1[,-15], train1$Absenteeism_time_in_hours, k = 1)
KNN_Predictions=as.numeric(as.character((KNN_Predictions)))
RMSE(test1[,15], KNN_Predictions)


####################Model Selection and Final Tuning#################################

##Linear Regression

lm_model = lm(Absenteeism_time_in_hours ~. , data = train1)
predictions_LR = predict(lm_model, test1[,-15])
RMSE(test1[,15], predictions_LR)

#Linear regression is giving good results compared to all algorithms. RMSE = 2.40 for linear regression

#Storing the predicted values in to CSV file
write.csv(predictions_LR, "model_predictions.csv", row.names = F)

########################Summary on the visualization ################################

## We observe that Reason for absence is the most important predictor in absenteeism hours
## Let us see the sum and mean of the absenteeism hours reason wise

reason_sum_hrs = aggregate(df$Absenteeism_time_in_hours,by = list(Category = df$Reason_for_absence),FUN = sum)
names(reason_sum_hrs)=c("Reason no.","Sum of absent hours")
reason_mean_hrs = aggregate(df$Absenteeism_time_in_hours,by = list(Category = df$Reason_for_absence),FUN = mean)
names(reason_mean_hrs)=c("Reason no.","Mean of absent hours")
table(df$Reason_for_absence)

########################## Monthly loss for the Company################################

loss_data = df[,c("Month_of_absence","Work_load_Average_per_day","Service_time","Absenteeism_time_in_hours")]
str(loss_data)
loss_data$WorkLoss = round((loss_data$Work_load_Average_per_day/loss_data$Service_time)*loss_data$Absenteeism_time_in_hours)
View(loss_data)
monthly_loss = aggregate(loss_data$WorkLoss,by = list(Category = loss_data$Month_of_absence),FUN = sum)
names(monthly_loss) = c("Month","WorkLoss")
ggplot(monthly_loss,aes(monthly_loss$Month,monthly_loss$WorkLoss))+geom_bar(stat = "identity",fill = "blue")+labs(y="WorkLoss",x="Months")

#################################################################################

#References:
#https://stackoverflow.com/questions/55270964/ggadjustedcurves-error-must-use-a-vector-in-not-an-object-of-class-matrix