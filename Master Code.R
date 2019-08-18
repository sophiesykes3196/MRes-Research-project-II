#Packages to install before running code 

install.packages("readxl")
install.packages("dpylr")
install.packages("XLConnect")
install.packages("XLConnectJars")
install.packages(("rJava"))
install.packages("xlsx")


####### DSSAT ##########

##STEP 1 - create DSSAT Sequence file 
library(readxl) # Read in packages required for DSSAT steps 
library(dplyr)

#Set working directory to where the master template is saved
workingdir<-"C:/Users/sophi/OneDrive/Documents/ZSL/Masterfile/Different scenarios"
setwd(workingdir)

# Reads in each sheet with DSSAT parameters from master template 
# and creates dataframes matching sequence file headings 

General<-read_xlsx("Master Sheet Two Scenarios.xlsx",1) 
General<- na.omit(General) # Removes empty rows 
Field.Harvest<-read_xlsx("Master Sheet Two Scenarios.xlsx",2)
Field<- data.frame(Field.Harvest[,1:24])
Field<- na.omit(Field)
Initial<- data.frame(Field.Harvest[,25:44])
Initial<- na.omit(Initial)
Harvest<-data.frame(Field.Harvest[,45:54])
Harvest<- na.omit(Harvest)
Fertiliser<-read_xlsx("Master Sheet Two Scenarios.xlsx",3)
Tillage.Res<-read_xlsx("Master Sheet Two Scenarios.xlsx",4)
Residues<-data.frame(Tillage.Res[,1:13])
Residues<- na.omit(Residues)
Tillage<- data.frame(Tillage.Res[,14:36])
Planting.Irr<-read_xlsx("Master Sheet Two Scenarios.xlsx",5)
Planting<- data.frame(Planting.Irr[,1:18])
Planting<- na.omit(Planting)
Irrigation<- data.frame(Planting.Irr[,19:33])
Irrigation<- na.omit(Irrigation)
Treatment.Sim<-read_xlsx("Master Sheet Two Scenarios.xlsx",6)
Treatments<- data.frame(Treatment.Sim[,1:19])
Treatments<- na.omit(Treatments)
Simulation<- data.frame (Treatment.Sim[,20:114])
Simulation<- na.omit(Simulation)
Simulation$SMODEL <- " "

## Extra step for fertilisers and tillage parameters- to transform data from horizontal to vertical format 

# Merge fertiliser columns into a dataframe
# Creates dataframe with all columns for fertiliser application 1 
fertiliserdf1 = data.frame(Fertiliser[,c("FDATE.1","FMCD.1","FACD.1","FDEP.1","FAMN.1","FAMP.1","FAMK.1","FAMC.1","FAMO.1","FOCD.1","Scenario")])
# Adds column names 
colnames(fertiliserdf1) <- c("FDATE","FMCD","FACD","FDEP","FAMN","FAMP","FAMK","FAMC","FAMO","FOCD","Scenario") 
fertiliserdf2 = data.frame(Fertiliser [,c("FDATE.2","FMCD.2","FACD.2","FDEP.2","FAMN.2","FAMP.2","FAMK.2","FAMC.2","FAMO.2","FOCD.2","Scenario")]) # Repeats for fertiliser application 2                              
colnames(fertiliserdf2) <- c("FDATE","FMCD","FACD","FDEP","FAMN","FAMP","FAMK","FAMC","FAMO","FOCD","Scenario")
fertiliserdf3 = data.frame(Fertiliser [,c("FDATE.3","FMCD.3","FACD.3","FDEP.3","FAMN.3","FAMP.3","FAMK.3","FAMC.3","FAMO.3","FOCD.3","Scenario")])# Repeats for fertiliser application 3 
colnames(fertiliserdf3) <- c("FDATE","FMCD","FACD","FDEP","FAMN","FAMP","FAMK","FAMC","FAMO","FOCD","Scenario")

# Adds @F and FERNAME column to other fertiliser dataframes as these are not in the master template 
fertiliserdf1$"@F"<- Fertiliser$X.F 
fertiliserdf1$FERNAME<-Fertiliser$FERNAME
fertiliserdf1= na.omit(fertiliserdf1) # removes any empty rows 
fertiliserdf2$"@F"<- Fertiliser$X.F #Repeats for fertiliser application 2 
fertiliserdf2$FERNAME<-Fertiliser$FERNAME
fertiliserdf2= na.omit(fertiliserdf2)
fertiliserdf3$"@F"<- Fertiliser$X.F #Repeats for fertiliser application 3 
fertiliserdf3$FERNAME<-Fertiliser$FERNAME
fertiliserdf3= na.omit(fertiliserdf3)

#Puts all three fertiliser application data into one dataframe which then can be read from in the write sequence file code
fertiliserdf= Reduce(function(x, y) merge(x, y, all=TRUE), list(fertiliserdf1, fertiliserdf2, fertiliserdf3))
fertiliserdf=na.omit(fertiliserdf) #remove any empty rows 
#Orders the fertiliser data by years
fertiliserdf= fertiliserdf[order(fertiliserdf$"@F"),]

# Make a dataframe with all tillage practices - same explanation as fertiliser df 
#Same explanation as fertiliser dataframe
tillagedf1 = data.frame (Tillage[,c("TDATE.1","TIMPL.1","TDEP.1")])
tillagedf1$Scenario = Tillage$Scenario...14
colnames(tillagedf1) <- c ("TDATE","TIMPL","TDEP","Scenario")
tillagedf1$"@T"<- Tillage$X.T
tillagedf1$TNAME<-Tillage$TNAME
tillagedf1= na.omit(tillagedf1)

tillagedf2 = data.frame (Tillage[,c("TDATE.2","TIMPL.2","TDEP.2")])
colnames(tillagedf2) <- c ("TDATE","TIMPL","TDEP")
tillagedf2$Scenario = Tillage$Scenario...14
colnames(tillagedf2) <- c ("TDATE","TIMPL","TDEP","Scenario")
tillagedf2$"@T"<- Tillage$X.T
tillagedf2$TNAME<-Tillage$TNAME
tillagedf2= na.omit(tillagedf2)

tillagedf3 = data.frame (Tillage[,c("TDATE.3","TIMPL.3","TDEP.3")])
colnames(tillagedf3) <- c ("TDATE","TIMPL","TDEP")
tillagedf3$Scenario = Tillage$Scenario...14
colnames(tillagedf3) <- c ("TDATE","TIMPL","TDEP","Scenario")
tillagedf3$"@T"<- Tillage$X.T
tillagedf3$TNAME<-Tillage$TNAME
tillagedf3= na.omit(tillagedf3)

tillagedf4 = data.frame (Tillage[,c("TDATE.4","TIMPL.4","TDEP.4")])
colnames(tillagedf4) <- c ("TDATE","TIMPL","TDEP")
tillagedf4$Scenario = Tillage$Scenario...14
colnames(tillagedf4) <- c ("TDATE","TIMPL","TDEP","Scenario")
tillagedf4$"@T"<- Tillage$X.T
tillagedf4$TNAME<-Tillage$TNAME
tillagedf4= na.omit(tillagedf4)

tillagedf5 = data.frame (Tillage[,c("TDATE.5","TIMPL.5","TDEP.5")])
colnames(tillagedf5) <- c ("TDATE","TIMPL","TDEP")
tillagedf5$Scenario = Tillage$Scenario...14
colnames(tillagedf5) <- c ("TDATE","TIMPL","TDEP","Scenario")
tillagedf5$"@T"<- Tillage$X.T
tillagedf5$TNAME<-Tillage$TNAME
tillagedf5= na.omit(tillagedf5)

tillagedf= Reduce(function(x, y) merge(x, y, all=TRUE), list(tillagedf1, tillagedf2, tillagedf3, tillagedf4, tillagedf5))
tillagedf= na.omit(tillagedf)
tillagedf=tillagedf[order(tillagedf$"@T"),]


##Change working directory to DSSAT sequence folder 
workingdir<-"C:/DSSAT47/Sequence"
setwd(workingdir)

## Create sequence function that writes a text file in the format of DSSAT's sequence file 
create_sqx_file <- function(file="temp.sqx", ex_details="IZBF0001", field_id="IZBF0001", site="01", scenario="WORST CASE SCENARIO",
                            General,Treatments,Field, Initial, Irrigation, Planting, Harvest, Simulation, tillagedf, fertiliserdf) {
  write(sprintf("*EXP.DETAILS: %s", ex_details), file=file)
  write(" ", file=file, append=T)
  write("*GENERAL", file=file, append=T)
  write("@PEOPLE", file=file, append=T)
  write("Alexa Varah", file=file, append=T)
  write("@ADDRESS", file=file, append=T)
  write("Institute of Zoology, Regents Park, London NW1 4RY", file=file, append=T)
  write("@SITE", file=file, append=T)
  write(sprintf("%s", site), file=file, append=T)
  
  write(" ", file=file, append=T)
  
  Treatments = subset(Treatments, Scenario...1 == scenario) # Subset for multiple scenarios- can be changed to subset field conditions 
  
  # Treatments
  write("*TREATMENTS                        -------------FACTOR LEVELS------------", file=file, append=T) 
  write("@N R O C TNAME.................... CU FL SA IC MP MI MF MR MC MT ME MH SM", file=file, append=T) #Writes columns headings in text file 
  
  ##A loop that pulls the data from each row of the master template 
  for (i in 1:nrow(Treatments)) {
    write(sprintf(" %-s%-2s %2s %-25s %s", Treatments[i,"X.N...2"], Treatments[i,"R"], paste(Treatments [i,c("O","C")], collapse=" "), Treatments[i,"TNAME"],
                  paste(apply(Treatments[i,c("CU","FL","SA","IC","MP","MI","MF","MR","MC","MT","ME","MH","SM")],
                              1,sprintf, fmt="%2s"), collapse = " ")), file=file, append=T)} 
  write(" ", file=file, append=T)
  
  
  General = subset(General, Scenario == scenario)
  
  # Cultivars
  
  write("*CULTIVARS", file=file, append=T)
  write("@C CR INGENO CNAME", file=file, append=T)
  for (i in 1:nrow(General)) {
    write(sprintf(" %2s", paste(General[i, c("C","CR","INGENO","CNAME")], collapse=" ")), file=file, append=T)
  }  
  
  write(" ", file=file, append=T)
  
  Field= subset(Field, Scenario...1 == scenario)
  
  # Fields
  write("*FIELDS", file=file, append=T)
  write("@L ID_FIELD WSTA....  FLSA  FLOB  FLDT  FLDD  FLDS  FLST SLTX  SLDP  ID_SOIL    FLNAME", file=file, append=T)
  
  #write(sprintf(" %s %8s %8s %s %4s %5s %11s %6s",
  for (i in 1:nrow(Field)){
    write(sprintf(" %-2s%8s %-8s %s %4s %5s %11s %6s",              
                  Field[i, "X.L"], Field[i,"ID_FIELD"], Field[i,"WSTA"],
                  paste(apply(Field[i, c("FLSA","FLOB","FLDT","FLDD","FLDS","FLST")], 1, sprintf, fmt="%5s"), collapse=" "),
                  Field[i,"SLTX"], Field[i, "SLDP"],
                  Field[i, "Soil.ID"], Field[i, "FLNAME"]), file=file, append=T) 
  }
  
  write("@L ...........XCRD ...........YCRD .....ELEV .............AREA .SLEN .FLWR .SLAS FLHST FHDUR", file=file, append=T)
  for (i in 1:nrow(Field)){write(sprintf(" %-2s%15s %15s %9s %17s %s", 
                                         Field[i, "X.L"], Field[i,"XCRD"], Field[i,"YCRD"], Field[i,"ELEV"], Field[i, "AREA"],
                                         paste(apply(Field[i, c(".SLEN",".FLWR",".SLAS","FLHST","FHDUR")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)}
  
  write(" ", file=file, append=T)
  
  # Initial conditions
  Initial = subset(Initial, Scenario...25 == scenario)
  
  write("*INITIAL CONDITIONS", file=file, append=T)
  write("@C   PCR ICDAT  ICRT  ICND  ICRN  ICRE  ICWD ICRES ICREN ICREP ICRIP ICRID ICNAME", file=file, append=T)
  for (i in 1:nrow(Initial)){write(sprintf(" %-2s%s %-6s", 
                                           Initial[i, "X.C"], 
                                           paste(apply(Initial[i,c("PCR","ICDAT","ICRT","ICND","ICRN","ICRE","ICWD","ICRES","ICREN","ICREP","ICRIP","ICRID")], 1, sprintf, fmt="%5s"), collapse=" "), 
                                           Initial[i, "ICNAME"]), file=file, append=T) }
  # Field initial conditions
  write("@C  ICBL  SH2O  SNH4  SNO3", file=file, append=T)
  for (i in 1:nrow(Initial)) {
    write(sprintf(" %-2s%s", 
                  Initial[i, "X.C" ], 
                  paste(apply(Initial[i, c("ICBL","SH2O","SNH4","SNO3")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
  }
  
  write(" ", file=file, append=T)
  
  #Planting details
  Planting= subset(Planting, Scenario...1 == scenario)
  write("*PLANTING DETAILS", file=file, append=T)
  write("@P PDATE EDATE  PPOP  PPOE  PLME  PLDS  PLRS  PLRD  PLDP  PLWT  PAGE  PENV  PLPH  SPRL                        PLNAME" , file=file, append=T)
  for (i in 1:nrow(Planting)) {
    write(sprintf(" %-2s%s %22s %-6s", 
                  Planting[i, "X.P"], 
                  paste(apply(Planting[i, c("PDATE","EDATE","PPOP","PPOE","PLME","PLDS","PLRS","PLRD","PLDP","PLWT","PAGE","PENV","PLPH","SPRL")], 1, sprintf, fmt="%5s"), collapse=" "), "", 
                  Planting[i, "PLNAME"]), file=file, append=T)
  }
  
  
  write(" ", file=file, append=T)
  
  # irrigation 
  #Irrigation= subset(Irrigation, Scenario...19 == scenario)
  write("*IRRIGATION AND WATER MANAGEMENT", file=file, append=T)
  write("@I  EFIR  IDEP  ITHR  IEPT  IOFF  IAME  IAMT IRNAME", file=file, append=T)
  for (i in 1:nrow(Irrigation)) {write(sprintf(" %-2s%s %-6s", 
                                               Irrigation[i, "X.I...20"], 
                                               paste(apply(Irrigation[i, c("EFIR","IDEP","ITHR","IEPT","IOFF","IAME","IAMT")], 1, sprintf, fmt="%5s"), collapse=" "), 
                                               Irrigation[i, "IRNAME"]), file=file, append=T)}
  
  write("@I IDATE  IROP IRVAL", file=file, append=T)
  for (i in 1:nrow(Irrigation)){write(sprintf(" %-2s%s", 
                                              Irrigation[i, "X.I...20"], 
                                              paste(apply(Irrigation[i, c("IDATE","IROP","IRVAL")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)}
  
  write(" ", file=file, append=T)
  
  #Fertilisers inorganic 
  fertiliserdf= subset(fertiliserdf, Scenario == scenario)
  write("*FERTILIZERS (INORGANIC)", file=file, append=T)
  write("@F FDATE  FMCD  FACD  FDEP  FAMN  FAMP  FAMK  FAMC  FAMO  FOCD FERNAME", file=file, append=T)
  for (i in 1:nrow(fertiliserdf)) {
    write(sprintf(" %-2s%s %-6s", 
                  fertiliserdf[i, "@F"], 
                  paste(apply(fertiliserdf[i, c("FDATE","FMCD","FACD","FDEP","FAMN","FAMP","FAMK","FAMC","FAMO","FOCD")], 1, sprintf, fmt="%5s"), collapse=" "), 
                  fertiliserdf[i, "FERNAME"]), file=file, append=T)
  }
  
  write(" ", file=file, append=T)
  
  # residues 
  Residues= subset(Residues, Scenario...1 == scenario)
  write("*RESIDUES AND ORGANIC FERTILIZER", file=file, append=T)
  write("@R RDATE  RCOD  RAMT  RESN  RESP  RESK  RINP  RDEP  RMET RENAME", file=file, append=T)
  for (i in 1:nrow(Residues)) {
    write(sprintf(" %-2s%s %-5s", 
                  Residues[i, "X.R"], 
                  paste(apply(Residues[i,c("RDATE","RCOD","RAMT","RESN","RESP","RESK","RINP","RDEP","RMET")], 1, sprintf, fmt="%5s"), collapse=" "),
                  Residues[i, "RENAME"]), file=file, append=T)
  }
  
  write(" ", file=file, append=T)
  
  
  # chem applications 
  # NB these commented out as I am not applying chemicals (leaching doesn't respond to chemical applicns)
  
  #write("*CHEMICAL APPLICATIONS", file=file, append=T)
  #write("@C CDATE CHCOD CHAMT  CHME CHDEP   CHT..CHNAME", file=file, append=T)
  #for (i in 1:nrow(master)) {
  #  write(sprintf(" %-2s%s  %-6s", 
  #                master[i, "X.C], 
  #                paste(apply(master[i, c("CDATE","CHCOD","CHAMT","CHME","cHDEP","CHT...")], 1, sprintf, fmt="%5s"), collapse=" "), 
  #                master[i, "CHNAME"]), file=file, append=T)
  #}
  
  #write(" ", file=file, append=T)
  
  # tillage and rotations 
  tillagedf = subset(tillagedf, Scenario == scenario)
  write("*TILLAGE AND ROTATIONS", file=file, append=T)
  write("@T TDATE TIMPL  TDEP TNAME", file=file, append=T)
  for (i in 1:nrow(tillagedf)) {
    write(sprintf(" %-2s%s %-6s", 
                  tillagedf[i, "@T"], 
                  paste(apply(tillagedf[i,c("TDATE","TIMPL","TDEP")], 1, sprintf, fmt="%5s"), collapse=" "), 
                  tillagedf[i, "TNAME"]), file=file, append=T)
  }
  write(" ", file=file, append=T)
  
  # harvest details 
  Harvest= subset(Harvest, Scenario...45 == scenario)
  write("*HARVEST DETAILS", file=file, append=T)
  write("@H HDATE  HSTG  HCOM HSIZE   HPC  HBPC HNAME", file=file, append=T)
  for (i in 1:nrow(Harvest)) {
    write(sprintf(" %-2s%s %-6s", 
                  Harvest[i, "X.H"],
                  paste(apply(Harvest[i, c("HDATE","HSTG","HCOM","HSIZE","HPC","HBPC")], 1, sprintf, fmt="%5s"), collapse=" "), 
                  Harvest[i, "HNAME"]), file=file, append=T)
  }
  
  write(" ", file=file, append=T)
  
  
  # simulation 
  Simulation= subset(Simulation, Scenario...20 == scenario)
  
  write("*SIMULATION CONTROLS", file=file, append=T)
  
  
  for (i in 1:nrow(Simulation)) {
    write("@N GENERAL     NYERS NREPS START SDATE RSEED SNAME.................... SMODEL", file=file, append=T)
    write(sprintf(" %-2s%-11s %s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "GENERAL"],
                  paste(apply(Simulation[i, c("NYERS","NREPS","START","SDATE","RSEED","SNAME")], 1, sprintf, fmt="%5s"), collapse=" "), 
                  Simulation[i, "SMODEL"]), file=file, append=T)
    
    write("@N OPTIONS     WATER NITRO SYMBI PHOSP POTAS DISES  CHEM  TILL   CO2", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "OPTIONS"],
                  paste(apply(Simulation[i, c("WATER","NITRO","SYMBI","PHOSP","POTAS","DISES","CHEM","TILL","CO2")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write("@N METHODS     WTHER INCON LIGHT EVAPO INFIL PHOTO HYDRO NSWIT MESOM MESEV MESOL", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "METHODS"],
                  paste(apply(Simulation[i, c("WTHER","INCON","LIGHT","EVAPO","INFIL","PHOTO","HYDRO","NSWIT","MESOM","MESEV","MESOL")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write("@N MANAGEMENT  PLANT IRRIG FERTI RESID HARVS", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "MANAGEMENT"],
                  paste(apply(Simulation[i, c("PLANT","IRRIG","FERTI","RESID","HARVS")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write("@N OUTPUTS     FNAME OVVEW SUMRY FROPT GROUT CAOUT WAOUT NIOUT MIOUT DIOUT VBOSE CHOUT OPOUT", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "OUTPUTS"],
                  paste(apply(Simulation[i, c("FNAME","OVVEW","SUMRY","FROPT","GROUT","CAOUT","WAOUT","NIOUT","MIOUT","DIOUT","VBOSE","CHOUT","OPOUT")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    
    write(" ", file=file, append=T) 
    
    write("@  AUTOMATIC MANAGEMENT", file=file, append=T)
    write("@N PLANTING    PFRST PLAST PH2OL PH2OU PH2OD PSTMX PSTMN", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "PLANTING"],
                  paste(apply(Simulation[i, c("PFRST","PLAST","PH2OL","PH2OU","PH2OD","PSTMX","PSTMN")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write("@N IRRIGATION  IMDEP ITHRL ITHRU IROFF IMETH IRAMT IREFF", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "IRRIGATION"],
                  paste(apply(Simulation[i, c("IMDEP","ITHRL","ITHRU","IROFF","IMETH","IRAMT","IREFF")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write("@N NITROGEN    NMDEP NMTHR NAMNT NCODE NAOFF", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "NITROGEN"],
                  paste(apply(Simulation[i, c("NMDEP","NMTHR","NAMNT","NCODE","NAOFF")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write("@N RESIDUES    RIPCN RTIME RIDEP", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "RESIDUES"],
                  paste(apply(Simulation[i, c("RIPCN","RTIME","RIDEP")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write("@N HARVEST     HFRST HLAST HPCNP HPCNR", file=file, append=T)
    write(sprintf(" %-2s%-11s %s", 
                  Simulation[i, "X.N...29"], Simulation[i, "HARVEST"],
                  paste(apply(Simulation[i, c("HFRST","HLAST","HPCNP","HPCNR")], 1, sprintf, fmt="%5s"), collapse=" ")), file=file, append=T)
    
    write(" ", file=file, append=T)
  }
  write(" ", file=file, append=T)
  write("", file=file, append=T)
}


## Loop that runs function for each scenario or field in master template 
#Extracts details for each scenario or field and removes any duplicates 
Details<- data.frame(General[,c(".SXQ.file.name","EXP.Details","Field.ID","Unique.2.digit.site.code","Scenario","No.Rotations","Experiment.number")])
Details<- Details[!duplicated(Details$EXP.Details),]

##Loop for create sequence file function 
for (i in 1:nrow(Details)) {
  #   
  filename = paste0(Details$.SXQ.file.name [i], ".SQX")
  this_details = Details$EXP.Details[i]
  this_field = Details$Field.ID[i]
  this_site = Details$Unique.2.digit.site.code [i]
  this_scenario = Details$Scenario[i]
  #   
  create_sqx_file(file=filename, ex_details=this_details, field_id=this_field, site=this_site, scenario=this_scenario,
                  General,Treatments,Field, Initial, Irrigation, Planting, Harvest, Simulation, tillagedf, fertiliserdf)
}

###STEP 2 - create DSSAT Batch file 

#Create batch file function that writes text file in the format of DSSAT batch file 

create_batch_file <- function(sqx_file="IZE29901", rotation=6, exp_no=1) {
  file = paste0(sqx_file, "_batch.v47")
  write("$BATCH(SEQUENCE)", file=file)
  write("!",  file=file, append=T)
  write("! Directory    : C:\\DSSAT47\\Sequence",  file=file, append=T)
  write(sprintf("! Command Line : C:\\DSSAT47\\DSCSM047.EXE Q %s", file),  file=file, append=T)
  write("! Crop         : Sequence",  file=file, append=T)       
  write(sprintf("! Experiment   : %s.SQX", sqx_file),  file=file, append=T)
  write(sprintf("! ExpNo        : %s", exp_no),  file=file, append=T)
  write(sprintf("! Debug        : C:\\DSSAT47\\DSCSM047.EXE \" Q %s\"", file),  file=file, append=T)
  write("!",  file=file, append=T)
  write("@FILEX                                                                                        TRTNO     RP     SQ     OP     CO",  file=file, append=T)
  for (i in 1:rotation) {
    write(sprintf("C:\\DSSAT47\\Sequence\\%s.SQX                                                                  1      1    %2d      1      0", sqx_file, i),  file=file, append=T)
  }
}

# Loop that runs function for each scenario or field in master template 
##Loop for create batch file function 

create_batch_file()
sqx_files = Details$.SXQ.file.name
exp_no = Details$Experiment.number
rotation = Details$No.Rotations
for (i in 1:length(sqx_files)) {
  create_batch_file(sqx_files[i], rotation[i], exp_no[i])
}

###STEP 3 - Open DSSAT and run crop simulation model for each sequence and batch file 

all_output = list()
sqx_files = Details$.SXQ.file.name

#Runs loop for each sequence file 
for (i in 1:length(sqx_files)) 
{
  exper = sqx_files[i]
  
  tryCatch(
    {
      batch_name = paste(exper, "_batch.v47", sep="") 
      print(sprintf("[%d of %d] Running: %s..", i, length(sqx_files), batch_name))
      # Opens DSSAT and pastes batch file 
      temp=system(paste('C:/DSSAT47/DSCSM047.EXE Q ', batch_name), intern=TRUE, wait=TRUE, invisible = FALSE)  #Call to run CSM within R
      
      #Having captured the summary information from the DSDAT run, save it to a file called *_summaryoutput.txt
      # get rid of empty lines
      temp = temp[lapply(temp, nchar) > 1]
      output_filename = paste(exper, "_summaryoutput.txt", sep="")
      write(temp, file=output_filename)
      
      # Work out .OSU (DSSAT Summary file) filename from experiment
      osu_filename = paste(exper, ".OSU", sep="")
      
      ## Pull out information about nitrate leaching
      #Open that file, skipping first three rows, and anything that starts with a star will be ignored (as a comment)
      leaching_data = read.table(file=osu_filename, skip=3, header = T, comment.char = "*", fill=T)
      colnames(leaching_data) = c(colnames(leaching_data)[2:ncol(leaching_data)], "NA") # exclude the first column name (it's an '@' character) and append an NA column at the end instead
      
      
      leaching_dataframe = data.frame(experiment=exper, field=leaching_data$FNAM, harvestdate=leaching_data$HDAT, crop=leaching_data$CR, nlcm=leaching_data$NLCM, hwah=leaching_data$HWAH) 
    }, error = function(e) {
      leaching_dataframe = data.frame(experiment=exper, field=NA, harvestdate=NA, crop=NA, nlcm=NA, hwah=NA) 
    })
  all_output[[i]] = leaching_dataframe
}

## Extract rotation sequence file name, rotation number, harvest date, harvest yield and nitrate leached and write into a summary .csv file 
library(data.table)
all_output_table = rbindlist(all_output)

write.csv(all_output_table, file = "all_output_table.csv")

##DSSAT crop simulation model completed 

##Clears all objects and packages used for DSSAT as some mask packages used for Cool Farm Tool
rm(list = ls(all = TRUE))
detach(package: readxl)

########## Cool Farm Tool #############

library(dplyr)
library("XLConnectJars")
#Can set java.parameters = "-Xmx1024m" as this allocates 2GB of RAM to rJava if using a large dataset
options(java.parameters = "-Xmx1024m")
#Need to run r 3.6.1. version for package to work 
library("XLConnect")

#Set working directory back to where the master template is saved 
workingdir <- "C:/Users/sophi/OneDrive/Documents/ZSL/Masterfile/Different scenarios"  # Sets working directory of files
setwd(workingdir)

#Loading CFT sheet from masterfile
CFTData = loadWorkbook("Mastet sheet Three scenarios.xlsx")
CFTData = readWorksheet(CFTData, sheet = "CFT")

#Load in Cool Farm Tool 
CFT <- XLConnect::loadWorkbook("CoolFarmTool.xlsm")  
setStyleAction(CFT,XLC$"STYLE_ACTION.NONE")  # Keeps the formatting of the original document

writeWorksheet(CFT," - United Kingdom", 2, 7, 5, header = FALSE)  # Writes the initial value, header = FALSE stops weird placement
writeWorksheet(CFT,"Metric", 2, 8, 5, header = FALSE)

#Create lookup table of where each column in CFT sheet is sent in Cool Farm Tool workbook 
lookup = data.frame(matrix( data = NA, nrow = 37, ncol = 3))  # Makes size of data frame of CFT data
lknames = c("DestSheet","DestCol","DestRow")
lookup <- setNames(lookup, lknames)
rnames = names(CFTData)
row.names(lookup) <- rnames

lookup$DestCol <- c(NA,NA,"e",NA,NA,"e","e","e","e","e","e","e","e","e","e",NA,"e","e","g","e","e","e","e","e","e","e","e","e","e","e","e","e","e","e","e","f","g") 
lookup$DestSheet <- as.numeric(c("NA","NA",2,"NA","NA",2,2,2,2,2,3,3,3,3,3,"NA",3,3,3,3,4,4,6,6,6,6,6,6,6,6,6,6,6,6,6,4,4))
lookup$DestRow <- as.numeric(c("NA","NA",6,"NA","NA",11,12,14,17,18,5,12,13,14,15,"NA",16,19,19,27,12,16,36,38,39,41,42,43,46,49,50,51,54,56,60,12,12))

myLetters <- letters[1:26]  # Function that matches letter to number for destination column
lookup$DestCol <- match(lookup$DestCol, myLetters)

# Create a function to run Cool Farm Tool with data in master template 

CFT_all <- function(DF, subtract = "nothing"){
  
  if(subtract == "herb fuel"){
    DF$applicationrate = 0
  }else if(subtract == "herb manuf"){
    DF$pesticide_applicns = 0
  }else if(subtract == "till"){
    DF[,23:34] = 0
  }else{
    DF = DF
  } ## Subset function if want to leave out any of the sources of emissions e.g. fertiliser use. 
  output = data.frame(fix.empty.names = FALSE)
  for (a in 1:nrow(DF)) {
    
    # For loop reads each field and puts it into a writable format for CFT
    
    X = DF[a,]  # Currently only reads one
    X <- t(X)  # Transposes Z to allow it to be merged with the lookup table
    
    
    FieldID = X[5,]
    print.default(a)
    X <- merge(X, lookup, by = "row.names")  # Gives destinations of each variable
    
    
    X <- X[!is.na(X$DestCol),]  # Removes unnecessary data
    X <- setNames(X, c("Type","Value","DestSheet","DestCol","DestRow"))
    
    for(i in 1:nrow(X)){
      
      # For loop reads each row value and writes it to the destination cell as found from the lookup table
      
      b = X$Value[i]
      c = X$DestSheet[i]
      d = X$DestRow[i]
      e = X$DestCol[i]
      writeWorksheet(CFT,b, c, d, e, header = FALSE)
    }
    
    writeWorksheet(CFT,"N", 3, 19, 6, header = FALSE)  # Sets fertiliser nutrient
    writeWorksheet(CFT,"Ammonium nitrate - 35% N", 3, 19, 5, header = FALSE)  # Sets fertiliser
    setForceFormulaRecalculation(CFT, sheet = 3, TRUE)  # Forces excel to re-evaluate all formulae
    
    results = readWorksheet(CFT,9,13,8,26,8) # Reads all results from results table per hectare
    results[is.na(results)] = 0
    results <- rbind(results, c(sum(results$Per.hectare))) 
    results <- setNames(results, "Value")
    results <- data.frame(c(FieldID,results$Value[1],results$Value[2],results$Value[4],results$Value[10],results$Value[14]),fix.empty.names = FALSE) 
    output <- rbind(output,t(results))
    
    
  }
  output <- setNames(output,c("Field ID","Fertiliser","N20","Pest","Till","Total")) ## Pulls out output from Cool Farm Tool 
  return(output)
}

#Runs function for all scenarios or fields in master template, writes summary .csv file with output 
CFTData$fertiliser1 = NA  # Fertiliser treated later

outputdata <- CFT_all(CFTData) #Ignore error messages 
write.csv(outputdata, file = "CFTsummary.csv") #Can change name of the file 






