library(readxl)
library(coin)
library(xts)
library(dygraphs)
library(ggplot2)

setwd("/home/alf/Scrivania/lav_hgts")


ao_sheets=excel_sheets("ao.xls")
nao_sheets=excel_sheets("nao.xls")

aoCOOLING_SUN_LOW=read_xls("ao.xls",ao_sheets[1])
aoCOOLING_SUN_LOW=as.data.frame(apply(aoCOOLING_SUN_LOW,2,as.numeric))

aoCOOLING_SUN_NEUTRAL=read_xls("ao.xls",ao_sheets[3])
aoCOOLING_SUN_NEUTRAL=as.data.frame(apply(aoCOOLING_SUN_NEUTRAL,2,as.numeric))

aoCOOLING_SUN_HIGH=read_xls("ao.xls",ao_sheets[2])
aoCOOLING_SUN_HIGH=as.data.frame(apply(aoCOOLING_SUN_HIGH,2,as.numeric))

aoSSW_SUN_LOW=read_xls("ao.xls",ao_sheets[4])
aoSSW_SUN_LOW=as.data.frame(apply(aoSSW_SUN_LOW,2,as.numeric))

aoSSW_SUN_NEUTRAL=read_xls("ao.xls",ao_sheets[6])
aoSSW_SUN_NEUTRAL=as.data.frame(apply(aoSSW_SUN_NEUTRAL,2,as.numeric))

aoSSW_SUN_HIGH=read_xls("ao.xls",ao_sheets[5])
aoSSW_SUN_HIGH=as.data.frame(apply(aoSSW_SUN_HIGH,2,as.numeric))

naoCOOLING_SUN_LOW=read_xls("nao.xls",nao_sheets[1])
naoCOOLING_SUN_NEUTRAL=read_xls("nao.xls",nao_sheets[3])
naoCOOLING_SUN_HIGH=read_xls("nao.xls",nao_sheets[2])
naoSSW_SUN_LOW=read_xls("nao.xls",nao_sheets[4])
naoSSW_SUN_NEUTRAL=read_xls("nao.xls",nao_sheets[6])
naoSSW_SUN_HIGH=read_xls("nao.xls",nao_sheets[5])

naoCOOLING_SUN_LOW=as.data.frame(apply(naoCOOLING_SUN_LOW,2,as.numeric))
naoCOOLING_SUN_NEUTRAL=as.data.frame(apply(naoCOOLING_SUN_NEUTRAL,2,as.numeric))
naoCOOLING_SUN_HIGH=as.data.frame(apply(naoCOOLING_SUN_HIGH,2,as.numeric))

naoSSW_SUN_LOW=as.data.frame(apply(naoSSW_SUN_LOW,2,as.numeric))
naoSSW_SUN_NEUTRAL=as.data.frame(apply(naoSSW_SUN_NEUTRAL,2,as.numeric))
naoSSW_SUN_HIGH=as.data.frame(apply(naoSSW_SUN_HIGH,2,as.numeric))

##########################################################################################
# ao_sheets
# [1] "COOLING_SUN_LOW"    
# [2] "COOLING_SUN_HIGH"   
# [3] "COOLING_SUN_NEUTRAL"
# [4] "SSW_SUN_LOW"        
# [5] "SSW_SUN_HIGH"       
# [6] "SSW_SUN_NEUTRAL"    
# nao_sheets
# [1] "COOLING_SUN_LOW"    
# [2] "COOLING_SUN_HIGH"   
# [3] "COOLING_SUN_NEUTRAL"
# [4] "SSW_SUN_LOW"        
# [5] "SSW_SUN_HIGH"       
# [6] "SSW_SUN_NEUTRAL"   


#########################################################################################

dates_aoCOOLING_SUN_LOW=gsub("_","-",gsub("D_","",names(aoCOOLING_SUN_LOW)))[-1]
aoCOOLING_SUN_LOW=aoCOOLING_SUN_LOW[4:21] #18 cases

dates_aoCOOLING_SUN_HIGH=gsub("_","-",gsub("D_","",names(aoCOOLING_SUN_HIGH)))[-1]
aoCOOLING_SUN_HIGH=aoCOOLING_SUN_HIGH[8:13] #6 cases

dates_aoCOOLING_SUN_NEUTRAL=gsub("_","-",gsub("D_","",names(aoCOOLING_SUN_NEUTRAL)))[-1]
aoCOOLING_SUN_NEUTRAL=aoCOOLING_SUN_NEUTRAL[3:15] #13 cases

dates_aoSSW_SUN_NEUTRAL=gsub("_","-",gsub("D_","",names(aoSSW_SUN_NEUTRAL)))[-1]
aoSSW_SUN_NEUTRAL=aoSSW_SUN_NEUTRAL[2:11] #12 cases

dates_aoSSW_SUN_LOW=gsub("_","-",gsub("D_","",names(aoSSW_SUN_LOW)))[-1]
aoSSW_SUN_LOW=aoSSW_SUN_LOW[2:12] #9 cases

dates_aoSSW_SUN_HIGH=gsub("_","-",gsub("D_","",names(aoSSW_SUN_HIGH)))[-1]
aoSSW_SUN_HIGH=aoSSW_SUN_HIGH[2:6] #5 cases

###############################

dates_naoCOOLING_SUN_LOW=gsub("_","-",gsub("D_","",names(naoCOOLING_SUN_LOW)))[-1]
naoCOOLING_SUN_LOW=naoCOOLING_SUN_LOW[2:19] #18 cases

dates_naoCOOLING_SUN_HIGH=gsub("_","-",gsub("D_","",names(naoCOOLING_SUN_HIGH)))[-1]
naoCOOLING_SUN_HIGH=naoCOOLING_SUN_HIGH[2:7] #6 cases

dates_naoCOOLING_SUN_NEUTRAL=gsub("_","-",gsub("D_","",names(naoCOOLING_SUN_NEUTRAL)))[-1]
naoCOOLING_SUN_NEUTRAL=naoCOOLING_SUN_NEUTRAL[2:14] #13 cases

dates_naoSSW_SUN_NEUTRAL=gsub("_","-",gsub("D_","",names(naoSSW_SUN_NEUTRAL)))[-1]
naoSSW_SUN_NEUTRAL=naoSSW_SUN_NEUTRAL[2:11] #10 cases

dates_naoSSW_SUN_LOW=gsub("_","-",gsub("D_","",names(naoSSW_SUN_LOW)))[-1]
naoSSW_SUN_LOW=naoSSW_SUN_LOW[2:11] #10 cases

dates_naoSSW_SUN_HIGH=gsub("_","-",gsub("D_","",names(naoSSW_SUN_HIGH)))[-1]
naoSSW_SUN_HIGH=naoSSW_SUN_HIGH[2:6] #5 cases

#########################################################################################
# Low vs  HIGH

res_test_AOLH=list()
res_test_NAOLH=list()
res_test_AOLN=list()
res_test_NAOLN=list()
res_test_AONH=list()
res_test_NAONH=list()
res_test_AOLN_H=list()
res_test_NAOLN_H=list()

for ( i in 1:60) {
  res_test_AOLH[[i]]=wilcox.test(as.numeric(aoCOOLING_SUN_LOW[i,]),as.numeric(aoCOOLING_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_NAOLH[[i]]=wilcox.test(as.numeric(naoCOOLING_SUN_LOW[i,]),as.numeric(naoCOOLING_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_AOLN[[i]]=wilcox.test(as.numeric(aoCOOLING_SUN_LOW[i,]),as.numeric(aoCOOLING_SUN_NEUTRAL[i,]), correct=FALSE)$p.value
  res_test_NAOLN[[i]]=wilcox.test(as.numeric(naoCOOLING_SUN_LOW[i,]),as.numeric(naoCOOLING_SUN_NEUTRAL[i,]), correct=FALSE)$p.value
  res_test_AONH[[i]]=wilcox.test(as.numeric(aoCOOLING_SUN_NEUTRAL[i,]),as.numeric(aoCOOLING_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_NAONH[[i]]=wilcox.test(as.numeric(naoCOOLING_SUN_NEUTRAL[i,]),as.numeric(naoCOOLING_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_AOLN_H[[i]]=wilcox.test(as.numeric(c(aoCOOLING_SUN_NEUTRAL[i,],aoCOOLING_SUN_LOW[i,])),as.numeric(aoCOOLING_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_NAOLN_H[[i]]=wilcox.test(as.numeric(c(aoCOOLING_SUN_NEUTRAL[i,],aoCOOLING_SUN_LOW[i,])),as.numeric(naoCOOLING_SUN_HIGH[i,]), correct=FALSE)$p.value
  
}

res_cooling=data.frame(IDday=1:60,
               LvsH_AO=unlist(res_test_AOLH),
               LvsH_NAO=unlist(res_test_NAOLH),
               LvsN_AO=unlist(res_test_AOLN),
               LvsN_NAO=unlist(res_test_NAOLN),
               NvsH_AO=unlist(res_test_AONH),
               NvsH_NAO=unlist(res_test_NAONH), 
               LNvsH_AO=unlist(res_test_AOLN_H),
               LNvsH_NAO=unlist(res_test_NAOLN_H)
               
)
 
plot(res_cooling$LvsH_AO)
plot(res_cooling$LvsH_NAO)
plot(res_cooling$LvsN_AO)
plot(res_cooling$LvsN_NAO)
plot(res_cooling$NvsH_AO)
plot(res_cooling$NvsH_NAO)
plot(res_cooling$LNvsH_AO)
plot(res_cooling$LNvsH_NAO)

file.remove("Tabella_strat_event.xls")
XLConnect::writeWorksheetToFile("Tabella_strat_event.xls",res_cooling,"cooling")

#########################################################################################
#########################################################################################
# Low vs  HIGH

res_test_AOLH=list()
res_test_NAOLH=list()
res_test_AOLN=list()
res_test_NAOLN=list()
res_test_AONH=list()
res_test_NAONH=list()
res_test_AOLN_H=list()
res_test_NAOLN_H=list()

for ( i in 1:60) {
  res_test_AOLH[[i]]=wilcox.test(as.numeric(aoSSW_SUN_LOW[i,]),as.numeric(aoSSW_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_NAOLH[[i]]=wilcox.test(as.numeric(naoSSW_SUN_LOW[i,]),as.numeric(naoSSW_SUN_HIGH[i,]), correct=FALSE)$p.value
  
  res_test_AOLN[[i]]=wilcox.test(as.numeric(aoSSW_SUN_LOW[i,]),as.numeric(aoSSW_SUN_NEUTRAL[i,]), correct=FALSE)$p.value
  res_test_NAOLN[[i]]=wilcox.test(as.numeric(naoSSW_SUN_LOW[i,]),as.numeric(naoSSW_SUN_NEUTRAL[i,]), correct=FALSE)$p.value

  res_test_AONH[[i]]=wilcox.test(as.numeric(aoSSW_SUN_NEUTRAL[i,]),as.numeric(aoSSW_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_NAONH[[i]]=wilcox.test(as.numeric(naoSSW_SUN_NEUTRAL[i,]),as.numeric(naoSSW_SUN_HIGH[i,]), correct=FALSE)$p.value
  
  res_test_AOLN_H[[i]]=wilcox.test(as.numeric(c(aoSSW_SUN_NEUTRAL[i,],aoSSW_SUN_LOW[i,])),as.numeric(aoSSW_SUN_HIGH[i,]), correct=FALSE)$p.value
  res_test_NAOLN_H[[i]]=wilcox.test(as.numeric(c(aoSSW_SUN_NEUTRAL[i,],aoSSW_SUN_LOW[i,])),as.numeric(naoSSW_SUN_HIGH[i,]), correct=FALSE)$p.value
  
}

res_warming=data.frame(IDday=1:60,
                       LvsH_AO=unlist(res_test_AOLH),
                       LvsH_NAO=unlist(res_test_NAOLH),
                       LvsN_AO=unlist(res_test_AOLN),
                       LvsN_NAO=unlist(res_test_NAOLN),
                       NvsH_AO=unlist(res_test_AONH),
                       NvsH_NAO=unlist(res_test_NAONH), 
                       LNvsH_AO=unlist(res_test_AOLN_H),
                       LNvsH_NAO=unlist(res_test_NAOLN_H)
)

plot(res_warming$LvsH_AO)
plot(res_warming$LvsH_NAO)
plot(res_warming$LvsN_AO)
plot(res_warming$LvsN_NAO)
plot(res_warming$NvsH_AO)
plot(res_warming$NvsH_NAO)


XLConnect::writeWorksheetToFile("Tabella_strat_event.xls",res_warming,"warming")



aoCOOLING_SUN_LOWm <-apply(aoCOOLING_SUN_LOW,1,mean,na.rm=T)
aoCOOLING_SUN_HIGHm<-apply(aoCOOLING_SUN_HIGH,1,mean,na.rm=T)
aoCOOLING_SUN_NEUTRALm<-apply(aoCOOLING_SUN_NEUTRAL,1,mean,na.rm=T)
aoSSW_SUN_NEUTRALm<-apply(aoSSW_SUN_NEUTRAL,1,mean,na.rm=T)
aoSSW_SUN_LOWm<-apply(aoSSW_SUN_LOW,1,mean,na.rm=T)
aoSSW_SUN_HIGHm<-apply(aoSSW_SUN_HIGH,1,mean,na.rm=T)

aoCOOLING_SUN_LOWsd<-apply(aoCOOLING_SUN_LOW,1,sd,na.rm=T)
aoCOOLING_SUN_HIGHsd<-apply(aoCOOLING_SUN_HIGH,1,sd,na.rm=T)
aoCOOLING_SUN_NEUTRALsd<-apply(aoCOOLING_SUN_NEUTRAL,1,sd,na.rm=T)
aoSSW_SUN_NEUTRALsd<-apply(aoSSW_SUN_NEUTRAL,1,sd,na.rm=T)
aoSSW_SUN_LOWsd<-apply(aoSSW_SUN_LOW,1,sd,na.rm=T)
aoSSW_SUN_HIGHsd<-apply(aoSSW_SUN_HIGH,1,sd,na.rm=T)
aoCOOLING_SUN_LNm<-apply(cbind(aoCOOLING_SUN_LOW,aoCOOLING_SUN_NEUTRAL),1,mean,na.rm=T)
aoCOOLING_SUN_LNsd<-apply(cbind(aoCOOLING_SUN_LOW,aoCOOLING_SUN_NEUTRAL),1,sd,na.rm=T)
naoCOOLING_SUN_LOWm <-apply(naoCOOLING_SUN_LOW,1,mean,na.rm=T)
naoCOOLING_SUN_HIGHm<-apply(naoCOOLING_SUN_HIGH,1,mean,na.rm=T)
naoCOOLING_SUN_NEUTRALm<-apply(naoCOOLING_SUN_NEUTRAL,1,mean,na.rm=T)
naoSSW_SUN_NEUTRALm<-apply(naoSSW_SUN_NEUTRAL,1,mean,na.rm=T)
naoSSW_SUN_LOWm<-apply(naoSSW_SUN_LOW,1,mean,na.rm=T)
naoSSW_SUN_HIGHm<-apply(naoSSW_SUN_HIGH,1,mean,na.rm=T)

naoCOOLING_SUN_LOWsd<-apply(naoCOOLING_SUN_LOW,1,sd,na.rm=T)
naoCOOLING_SUN_HIGHsd<-apply(naoCOOLING_SUN_HIGH,1,sd,na.rm=T)
naoCOOLING_SUN_NEUTRALsd<-apply(naoCOOLING_SUN_NEUTRAL,1,sd,na.rm=T)
naoSSW_SUN_NEUTRALsd<-apply(naoSSW_SUN_NEUTRAL,1,sd,na.rm=T)
naoSSW_SUN_LOWsd<-apply(naoSSW_SUN_LOW,1,sd,na.rm=T)
naoSSW_SUN_HIGHsd<-apply(naoSSW_SUN_HIGH,1,sd,na.rm=T)

naoCOOLING_SUN_LNm<-apply(cbind(naoCOOLING_SUN_LOW,naoCOOLING_SUN_NEUTRAL),1,mean,na.rm=T)
naoCOOLING_SUN_LNsd<-apply(cbind(naoCOOLING_SUN_LOW,naoCOOLING_SUN_NEUTRAL),1,sd,na.rm=T)

naoSSW_SUN_LNm<-apply(cbind(naoSSW_SUN_LOW,naoSSW_SUN_NEUTRAL),1,mean,na.rm=T)
naoSSW_SUN_LNsd<-apply(cbind(naoSSW_SUN_LOW,naoSSW_SUN_NEUTRAL),1,sd,na.rm=T)

aoSSW_SUN_LNm<-apply(cbind(aoSSW_SUN_LOW,aoSSW_SUN_NEUTRAL),1,mean,na.rm=T)
aoSSW_SUN_LNsd<-apply(cbind(aoSSW_SUN_LOW,aoSSW_SUN_NEUTRAL),1,sd,na.rm=T)


res_data_df=data.frame(IDday=1:60,
                     aoCOOLING_SUN_LOWm,
                     aoCOOLING_SUN_LOWsd,
                     aoCOOLING_SUN_HIGHm,
                     aoCOOLING_SUN_HIGHsd,
                     aoCOOLING_SUN_NEUTRALm,
                     aoCOOLING_SUN_NEUTRALsd,
                     aoCOOLING_SUN_LNm,
                     aoCOOLING_SUN_LNsd,
                     aoSSW_SUN_NEUTRALm,
                     aoSSW_SUN_NEUTRALsd,
                     aoSSW_SUN_LOWm,
                     aoSSW_SUN_LOWsd,
                     aoSSW_SUN_HIGHm,
                     aoSSW_SUN_HIGHsd,
                     naoCOOLING_SUN_LOWm,
                     naoCOOLING_SUN_LOWsd,
                     naoCOOLING_SUN_HIGHm,
                     naoCOOLING_SUN_HIGHsd,
                     naoCOOLING_SUN_NEUTRALm,
                     naoCOOLING_SUN_NEUTRALsd,
                     naoSSW_SUN_NEUTRALm,
                     naoSSW_SUN_NEUTRALsd,
                     naoSSW_SUN_LOWm,
                     naoSSW_SUN_LOWsd,
                     naoSSW_SUN_HIGHm,
                     naoSSW_SUN_HIGHsd,
                     naoSSW_SUN_LNm,
                     naoSSW_SUN_LNsd,
                     naoCOOLING_SUN_LNm,
                     naoCOOLING_SUN_LNsd,
                     aoSSW_SUN_LNm,
                     aoSSW_SUN_LNsd)

saveRDS(res_data_df,"res_data_df.rds")
file.remove("Tabella_strat_stats.xls")
XLConnect::writeWorksheetToFile("Tabella_strat_stats.xls",res_data_df,"data")


res_data_df=readRDS("res_data_df.rds")
########################################################################################################

ggplot(res_data_df, aes(x = IDday, y = naoCOOLING_SUN_LNm)) + geom_point() +  geom_smooth(method = 'loess',color="green",linetype="dashed")+
  geom_point(color="green") +  geom_smooth(aes(x = IDday, y = naoCOOLING_SUN_HIGHm), method = 'loess',color="red",linetype="dashed")+
  geom_point(aes(x = IDday, y = naoCOOLING_SUN_HIGHm),color="red") + 
  annotate("rect", xmin = 42,xmax = 49, ymin = -0.2, ymax = 2,fill = "red",alpha = .3) +
  annotate("text", x = 40, y =1.6, label = "NAO daily differences \n not null from 42 to 49 days \n( WT p.value < .05)") + 
  xlim(0, 60)+ ggtitle("Mean NAO Profile in the post-event 60 days:\n High (red) and Low/Neutral Solar Activity  (green)",subtitle="Strat Cooling Events") +
  xlab("days") + ylab("NAO index")

ggsave("nao_cooling.png")
########################################################################################################

ggplot(res_data_df, aes(x = IDday, y = aoCOOLING_SUN_LNm)) + geom_point() +  geom_smooth(method = 'loess',color="green",linetype="dashed")+
  geom_point(color="green") +  geom_smooth(aes(x = IDday, y = aoCOOLING_SUN_HIGHm),method = 'loess',color="red",linetype="dashed")+ 
  geom_point(aes(x = IDday, y = aoCOOLING_SUN_HIGHm),color="red") + 
  annotate("rect", xmin = 39,xmax = 56, ymin = -1, ymax = 5.5,fill = "red",alpha = .3) +
  annotate("text", x = 40, y = 4.5, label = "AO/NAM-1000hPa in time daily differences \n not null from 39 to 56 days \n( WT p.value < .05)") + 
  xlim(0, 60)+ ggtitle("Mean AO/NAM-1000hPa Profile in the post-event 60\n days: High (red) and Low/Neutral Solar Activity (green)",subtitle="Strat Cooling Events") +
  xlab("days") + ylab("AO index")
ggsave("ao_cooling.png")

########################################################################################################


ggplot(res_data_df, aes(x = IDday, y = naoSSW_SUN_LNm)) + geom_point() +  geom_smooth(method = 'loess', color="green",linetype="dashed")+
  geom_point(color="green") +  geom_smooth(aes(x = IDday, y = naoSSW_SUN_HIGHm),method = 'loess', color="red",linetype="dashed")+
  geom_point(aes(x = IDday, y = naoSSW_SUN_HIGHm),color="red") + 
  annotate("text", x = 48, y =0.4, label = "No significant NAO time\n daily differences  \n( WT p.value < .05)") + 
  xlim(0, 60)+ ggtitle("Mean NAO Profile in the post-event 60 days:\n High (red) and Low/Neutral Solar Activity State (green)",subtitle="Strat Warming Events") +
  xlab("days") + ylab("NAO index")
ggsave("nao_warming.png")

########################################################################################################

ggplot(res_data_df, aes(x = IDday, y = aoSSW_SUN_LNm)) + geom_point() +  geom_smooth(method = 'loess', color="green", linetype="dashed")+
  geom_point(color="green") +  geom_smooth(aes(x = IDday, y = aoSSW_SUN_HIGHm),method = 'loess', color="red",linetype="dashed")+
  geom_point(aes(x = IDday, y = aoSSW_SUN_HIGHm), color="red") + 
  annotate("text", x = 40, y =0.4, label = "No significant AO/NAM-1000hPa\n time daily differences (WT p.value < .05)") + 
  xlim(0, 60)+ ggtitle("Mean  AO/NAM-1000hPa Profile in the post-event 60 days:\n High (red) and Low/Neutral Solar Activity (green)",subtitle="Strat Warming Events") +
  xlab("days") + ylab("AO index")

ggsave("ao_warming.png")

