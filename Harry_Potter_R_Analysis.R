library(readxl)

HP1<- read_excel("C:/Users/OR/PycharmProjects/CoffeeAndHappiness/LinesByGenderHP1.xlsx")
HP2<- read_excel("C:/Users/OR/PycharmProjects/CoffeeAndHappiness/LinesByGenderHP2.xlsx")
HP3<- read_excel("C:/Users/OR/PycharmProjects/CoffeeAndHappiness/LinesByGenderHP3.xlsx")
HP4<- read_excel("C:/Users/OR/PycharmProjects/CoffeeAndHappiness/LinesByGenderHP4.xlsx")
HP6<- read_excel("C:/Users/OR/PycharmProjects/CoffeeAndHappiness/LinesByGenderHP6.xlsx")


############################# Data Extraction #############################

##############
#### HP1 #####
##############

HP1Female<- HP1[c(1:7),3]
HP1Male<- HP1[c(8:33),3]


##############
#### HP2 #####
##############

HP2Female<- HP2[c(1:8),3]
HP2Male<- HP2[c(9:35),3]


##############
#### HP3 #####
##############

HP3Female<- HP3[c(1:11),3]
HP3Male<- HP3[c(12:35),3]


##############
#### HP4 #####
##############

HP4Female<- HP4[c(1:9),3]
HP4Male<- HP4[c(10:31),3]


##############
#### HP6 #####
##############

HP6Female<- HP6[c(1:9),3]
HP6Male<- HP6[c(10:28),3]


############################# Data Analysis #############################

print("************************* Data Analysis *************************")

HPFemale<-c(HP1Female, HP2Female, HP3Female, HP4Female, HP6Female)
HPMale<-c(HP1Male, HP2Male, HP3Male, HP4Male, HP6Male)

for(i in 1:(length(HPFemale)-1))
{
  cat("######### Harry Potter", i, "#########")
  cat(sep="\n\n")
  cat(sep="\n\n")
  cat("Variances are Equal in Harry Potter", i)
  cat(sep="\n\n")
  print(var.test(unlist(HPFemale[i]), unlist(HPMale[i]), alternative = "two.sided", sep="\n\n"))
  cat(sep="\n\n")
  cat("Female and Male characters have equal amount of lines in Harry Potter", i)
  cat(sep="\n\n")
  print(t.test(unlist(HPMale[i]), unlist(HPFemale[i]), mu=0 ,paired = FALSE, conf.level = 0.95))
  cat(sep="\n\n")
  cat(sep="\n\n")
}
cat("######### Harry Potter 6 #########", sep="\n\n")
cat("Variances Differ in Harry Potter 6")
var.test(unlist(HP6Female), unlist(HP6Male), alternative = "two.sided")
