# set working directory first

library(agricolae)
library(writexl)
library(dplyr)
library(data.table)

# import the example excel sheet 1 first as named here "sheet"

#CHANGE (character name)  below
#CHANGE (character name)  below
model.lm <- lm(sheet$`Character 1`~sheet$Cultivar)
model.aov <- aov(sheet$`Character 1`~sheet$Cultivar)
totalstats <- model.tables(model.aov, type = "means", se = TRUE)
anova <- anova(model.lm)
anova

#CHANGE (character name)  below
#CHANGE (DFerror = 'from Residual's DF value)  below
#CHANGE (MSerror = 'from Residual's Mean Sq value) below
dmrt <- duncan.test(sheet$`Character 1`, sheet$Cultivar, alpha = 0.05,
                    DFerror = 22, MSerror = 2.0936)

st.error <- c(totalstats$se)  
grandmean <- c(totalstats$tables$`Grand mean`)
df <- data.frame(as.data.table(dmrt$groups, keep.rownames = TRUE))
df1 <- data.frame(as.data.table(dmrt$statistics, keep.rownames = TRUE))
df2 <- data.frame(as.data.table(dmrt$parameters, keep.rownames = TRUE))
df3 <- data.frame(as.data.table(dmrt$duncan, keep.rownames = TRUE))
df4 <- data.frame(as.data.table(dmrt$means, keep.rownames = TRUE))
df5 <- data.frame(as.data.table(dmrt$comparison, keep.rownames = TRUE))
df6 <- data.frame(as.data.table(anova, keep.rownames = TRUE))
df7 <- data.frame(st.error, grandmean)

finalxlsheet <- list ('Groups' = df, 'Means' = df4, 'StandardError' = df7, 
                   'anova' = df6, 'Stats' = df1, 'Parameters' = df2,
                   'Duncan' = df3, 'comparison' = df5)

#CHANGE (character name)  below
write_xlsx(finalxlsheet, 'Character 1.xlsx')

# in the excel sheet you got 'rn' refer to rownames & 'X3' refers  Standard error
