library(readxl)
library(xlsx)
library(ggplot2)

results = data.frame(year= character(0),capacity_MW= character(0),robustness=character(0),avg_price_change = character(0),avg_org_nor_price= character(0),avg_ger_price = character(0),avg_new_nor_price = character(0),avg_ger_nor_price_diff= character(0),avg_ABS_ger_nor_price_diff= character(0),avg_perc_diff_weighted= character(0),avg_percent_diff_nonweighted= character(0),sum_price_sell_change= character(0),sum_price_buy_change= character(0),sum_PS_change= character(0),sum_CS_change= character(0),sum_welfare_change= character(0),country=character(0))
#creates a dataframe for final results 
country = "UK" #"germany" or "UK"
for (year in 2015:2018)
  
{ 
  
  foreign_prices = read_excel(paste("~/Desktop/electricity/data/",country,"_prices/Day-ahead_Prices_",year,".xlsx",sep = ""), col_names = FALSE)
  #loads the price data for foreign country
  
  for (capacity_MW in c(700,1400,2100,2800)) 
  {#set the transmission capacity
    
    
    for (robustness in c(1,0.8,0.6))
    {#set the parameter for robustness check
      
      
      intersections = data.frame(time = character(0),original_price = numeric(0),original_quantity = numeric(0),german_price = numeric(0), new_price = numeric(0), new_quantity = numeric(0), ger_nor_org_diff = numeric(0), price_change = numeric(0), perc_diff = numeric(0), price_buy_change= numeric(0),CS_change= numeric(0),price_sell_change= numeric(0),PS_change= numeric(0),welfare_change= numeric(0))
      #creates a data frame for intersections
      
      
      data_files = list.files(paste("~/Desktop/electricity/data/",year,"_curves",sep=""))
      for (data_file in data_files)
      { 
        
        ################################################################
        data_copy = read_excel(paste("~/Desktop/electricity/data/",year,"_curves/",data_file, sep=""), col_names = FALSE) 
        #data_copy = read_excel("~/Desktop/electricity/data/2018_curves/2018-10-28.xlsm",col_names = FALSE)
        ################################################################
        
        
        for (hour in 1:24) 
          
        {
          #extracting buy curve data:
          hour_buy = data.frame(data_copy[((which(data_copy[, hour * 2 - 1] == "Buy curve"))+1):((which(data_copy[, hour * 2 - 1] == "Sell curve"))-1), hour * 2] )
          #extracts the buy data in every second column from row 15 to the row prior to "Sell curve" in respective column (every second column -1)
          names(hour_buy) = c("column")
          #set the name of the only column in hour_buy dataframe to "column" so that it is uniform 
          hour_buy_price = hour_buy$column[ c(T,F) ]
          #extracts every first value from the hour_buy dataframe which is the price value
          hour_buy_price = as.numeric(hour_buy_price)
          #converts the character values to numeric
          hour_buy_q = hour_buy$column[ c(F,T) ]
          #extracts every second value from the hour_buy dataframe which is the quantity value
          hour_buy_q = as.numeric(hour_buy_q)
          #converts the character values to numeric
          hour_buy_curve = data.frame(hour_buy_price, hour_buy_q, stringsAsFactors = FALSE)
          # makes the dataframe of quantity and price for the buy curve
          
          #extracting sell curve data:
          hour_sell = data.frame(data_copy[((which(data_copy[, hour * 2 - 1] == "Sell curve"))+1):nrow(data_copy), hour * 2] )
          # extracts the sell data in every second column from row of "Sell curve" to the end of dataframe
          names(hour_sell) = c("column")
          #set the name of the only column in hour_sell dataframe to "column" so that it is uniform 
          hour_sell_price = hour_sell$column[ c(T,F) ]
          #extracts every first value from the hour_sell dataframe which is the price value
          hour_sell_price <- as.numeric(hour_sell_price)
          #converts the character values to numeric
          hour_sell_q = hour_sell$column[ c(F,T) ]
          #extracts every second value from the hour_sell dataframe which is the quantity value
          hour_sell_q = as.numeric(hour_sell_q)
          #converts the character values to numeric
          hour_sell_curve <- data.frame(hour_sell_price, hour_sell_q, stringsAsFactors = FALSE)
          #makes the dataframe of quantity and price for the sell curve
          hour_sell_curve = na.omit(hour_sell_curve)
          #deletes the N/A values from the dataframe
          
          #adjusting the curves with respect to block orders and import/export:
          hour_sell_curve[,2] = hour_sell_curve[,2]+as.numeric(data_copy[5,hour*2])
          #adding accepted blocks sell
          hour_buy_curve[,2] = hour_buy_curve[,2]+as.numeric(data_copy[4,hour*2])
          #adding accepted blocks buy
          if(as.numeric(data_copy[(which(data_copy[, hour * 2 - 1] == "Bid curve chart data (Volume for net flows)")),hour*2])>0) {hour_sell_curve[,2] = hour_sell_curve[,2]+as.numeric(data_copy[(which(data_copy[, hour * 2 - 1] == "Bid curve chart data (Volume for net flows)")),hour*2])}
          #if net flow >0 (import) - shift sell curve right (increased quantity and decreased price)
          if(as.numeric(data_copy[(which(data_copy[, hour * 2 - 1] == "Bid curve chart data (Volume for net flows)")),hour*2])<0) {hour_buy_curve[,2] = hour_buy_curve[,2]-as.numeric(data_copy[(which(data_copy[, hour * 2 - 1] == "Bid curve chart data (Volume for net flows)")),hour*2])}
          #if net flow <0 (export) - shift buy curve right (increaed quantity and increased price)
          
          #script for finding the intersection of curves taken from: https://rdrr.io/github/andrewheiss/reconPlots/src/R/curve_intersect.R - eddited
          curve_intersect = function(curve1, curve2) {
            {
              # Approximate the functional form of both curves:
              curve1_f <- approxfun(curve1[,1], curve1[,2], rule = 2)
              curve2_f <- approxfun(curve2[,1], curve2[,2], rule = 2)
              # Calculate the intersection of curve 1 and curve 2 along the x-axis:
              point_x <- uniroot(function(x) curve1_f(x) - curve2_f(x),
                                 c(min(curve1[,1]), max(curve1[,1])))$root
              # Find where point_x is in curve 2:
              point_y <- curve2_f(point_x)
            } 
            return(data.frame(x = point_x, y = point_y))
          }
          
          intersection = curve_intersect(hour_buy_curve, hour_sell_curve)
          #comand for finding the intersection (using the script above)
          
          if(country == "germany"){
            intersection = cbind(data_copy[1,hour*2],intersection,as.numeric(foreign_prices[(which(foreign_prices[, 1] == substr(data_copy[1,2], 1, 10)))+1+hour,2])) 
            #adds the time info and german price at that time to the intersection result
          }else if(country == "UK"){
            intersection = cbind(data_copy[1,hour*2],intersection,as.numeric(foreign_prices[(which(foreign_prices[, 6] == substr(data_copy[1,hour*2], 1, 13))),4])) 
            #adds the time info and UK price at that time to the intersection result
          }
          shift_hour_sell_curve = hour_sell_curve
          shift_hour_buy_curve = hour_buy_curve
          #copies curves which will be shiftet due to additional export/import 
          
          if(intersection[1,2]>intersection[1,4]) {shift_hour_sell_curve[,2]=hour_sell_curve[,2]+capacity_MW}
          #if Norwegian price larger tnan foreign price, then import capacity in MW (shift sell curve right)
          if(intersection[1,2]<intersection[1,4]) {shift_hour_buy_curve[,2]=hour_buy_curve[,2]+capacity_MW}
          #if Norwegian price smaller tnan foreign price, then export capacity in MW (shift buy curve right)
          
          new_intersection = curve_intersect(shift_hour_buy_curve, shift_hour_sell_curve)
          #comand for finding the intersection of shifted curve 
          
          intersection = cbind(intersection, new_intersection)
          #adds the new intersection into the result
          
          
          if(intersection[1,2]>intersection[1,4] && intersection[1,5]<intersection[1,4]) {intersection[1,5] = intersection[1,4]}
          #if original Norwegian price larger tnan foreign price and  New Norwegian price smaller than foreign price, then the new Norwegian price stops at foreign price
          
          if(intersection[1,2]<intersection[1,4] && intersection[1,5]>intersection[1,4]) {intersection[1,5] = intersection[1,4]}
          #if original Norwegian price smaller tnan foreign price and  New Norwegian price larger than foreign price, then the new Norwegian price stops at foreign price
          
          ger_nor_org_difference = intersection[1,4]-intersection[1,2]
          #calculates the difference in original prices
          
          price_change = intersection[1,5]-intersection[1,2]
          #clculates price change (new-old)
          
          perc_diff = 100/ger_nor_org_difference*price_change 
          #calculates how much % of the foreign-norwegian price difference is the price change
          
          if(perc_diff/100 > robustness){
            intersection[,5] = intersection[,2] + robustness*ger_nor_org_difference # correction of the error, price_change corrected to ger_nor_org_difference
            #new price = orgiginal price + robustness * price change
            #does not influence the volume, only price and subsequent calculations
          }
          
          #recalculation price change and perc_diff
          price_change = intersection[1,5]-intersection[1,2]
          perc_diff = 100/ger_nor_org_difference*price_change
          
          
          intersection = cbind(intersection, ger_nor_org_difference, price_change, perc_diff)
          #adds the data to the result
          
          #EXPORT WELFARE CALCULATION:
          if(intersection[1,2]<intersection[1,4]) { #original price lower than foreign (export)
            #Producer's surpluss change:
            sub_sell_curve = hour_sell_curve[which(abs(hour_sell_curve-intersection[,2])==min(abs(hour_sell_curve-intersection[,2]))):which(abs(hour_sell_curve-intersection[,5])==min(abs(hour_sell_curve-intersection[,5]))), ]
            #creates a data frame of the sell curve just between original and new price (approximately, correction follows) 
            sub_sell_curve = sub_sell_curve[seq(dim(sub_sell_curve)[1],1),]
            #reverse the order in sub curve such that the difference in next step makes sense (otherwise conflict between calculating upper and lower integral for buy or sell) (order is reversed back afterwards)
            difference=data.frame(-diff(as.matrix(sub_sell_curve$hour_sell_price)))
            #calculates the quantity increase for each price
            difference = rbind(difference,0)
            #adds zero to the last missing value which was caused by diff function so that columns can be merged
            sub_sell_curve = cbind(sub_sell_curve,difference)
            sub_sell_curve = sub_sell_curve[seq(dim(sub_sell_curve)[1],1),]
            #reverse the order of sub curve again to get to original order
            PS = sub_sell_curve[,2]*sub_sell_curve[,3]
            #calculates the change of producer surplus
            sub_sell_curve = cbind(sub_sell_curve,PS)
            #ads the PS to the sub_sell_curve dataframe
            price_sell_change = sum(sub_sell_curve[,3])
            PS_change = sum(sub_sell_curve[,4])
            
            #Consummer's surpluss change:
            sub_buy_curve = hour_buy_curve[which(abs(hour_buy_curve-intersection[,2])==min(abs(hour_buy_curve-intersection[,2]))):which(abs(hour_buy_curve-intersection[,5])==min(abs(hour_buy_curve-intersection[,5]))), ]
            #creates a data frame of the buy curve just between original and new price   (approximately, correction follows)
            #sub_buy_curve = sub_buy_curve[seq(dim(sub_buy_curve)[1],1),]
            #reverse the order in sub curve such that the difference in next step makes sense (order is reversed back afterwards)
            difference=data.frame(diff(as.matrix(sub_buy_curve$hour_buy_price)))
            #calculates the quantity increase for each price
            difference = rbind(difference,0)
            #adds zero to the last missing value which was caused by diff function so that columns can be merged
            sub_buy_curve = cbind(sub_buy_curve,-difference)
            #sub_buy_curve = sub_buy_curve[seq(dim(sub_buy_curve)[1],1),]
            #reverse the order of sub curve again to get to original order
            CS = sub_buy_curve[,2]*sub_buy_curve[,3]
            #calculates the change of producer surpluss
            sub_buy_curve = cbind(sub_buy_curve,CS)
            #ads the PS to the sub_buy_curve dataframe
            price_buy_change = sum(sub_buy_curve[,3])
            CS_change = sum(sub_buy_curve[,4])
            
            #Not used part of code:
            #if(abs(price_buy_change) < abs(price_sell_change)){
            #CS_change = CS_change-(abs(abs(price_buy_change) - abs(price_sell_change)))*(mean(hour_sell_curve$hour_sell_q))
            #}
            #if(abs(price_sell_change) < abs(price_buy_change)){
            #PS_change = PS_change+(abs(abs(price_buy_change) - abs(price_sell_change)))*(mean(hour_buy_curve$hour_buy_q))
            #}
            #welfare_change = PS_change+CS_change
          }
          
          #IMPORT WELFARE CALCULATION:
          if(intersection[1,2]>intersection[1,4]) { #original price larger than foreign (import)
            #consumer's surplus change
            sub_buy_curve = hour_buy_curve[which(abs(hour_buy_curve-intersection[,5])==min(abs(hour_buy_curve-intersection[,5]))):which(abs(hour_buy_curve-intersection[,2])==min(abs(hour_buy_curve-intersection[,2]))), ]
            #creates a data frame of the buy curve just between original and new price  (approximately, correction follows) 
            difference=data.frame(diff(as.matrix(sub_buy_curve$hour_buy_price)))
            #calculates the quantity increase for each price
            difference = rbind(difference,0)
            #adds zero to the last missing value which was caused by diff function so that columns can be merged
            sub_buy_curve = cbind(sub_buy_curve,difference)
            CS = sub_buy_curve[,2]*sub_buy_curve[,3]
            #calculates the change of CS
            sub_buy_curve = cbind(sub_buy_curve,CS)
            #ads the CS to the sub_buy_curve dataframe
            price_buy_change = sum(sub_buy_curve[,3])
            CS_change = sum(sub_buy_curve[,4])
            
            #producer's surplus change
            sub_sell_curve = hour_sell_curve[which(abs(hour_sell_curve-intersection[,5])==min(abs(hour_sell_curve-intersection[,5]))):which(abs(hour_sell_curve-intersection[,2])==min(abs(hour_sell_curve-intersection[,2]))), ]
            #creates a data frame of the sell curve just between original and new price   (approximately, correction follows)
            sub_sell_curve = sub_sell_curve[seq(dim(sub_sell_curve)[1],1),]
            #reverse the order in sub curve such that the difference in next step makes sense (otherwise conflict between calculating upper and lower integral for buy or sell) (order is reversed back afterwards)
            difference=data.frame(-diff(as.matrix(sub_sell_curve$hour_sell_price)))
            #calculates the quantity increase for each price
            difference = rbind(difference,0)
            #adds zero to the last missing value which was caused by diff function so that columns can be merged
            sub_sell_curve = cbind(sub_sell_curve,-difference)
            sub_sell_curve = sub_sell_curve[seq(dim(sub_sell_curve)[1],1),]
            #reverse the order of sub curve again to get to original order
            PS = sub_sell_curve[,2]*sub_sell_curve[,3]
            #calculates the change of PS
            sub_sell_curve = cbind(sub_sell_curve,PS)
            #ads the PS to the sub_sell_curve dataframe
            price_sell_change = sum(sub_sell_curve[,3])
            PS_change = sum(sub_sell_curve[,4])
            
            #correction for inelastic curves 
            #if(abs(price_buy_change) < abs(price_sell_change)){
            #CS_change = CS_change+(abs(abs(price_buy_change) - abs(price_sell_change)))*(mean(hour_sell_curve$hour_sell_q))
            #}
            #if(abs(price_sell_change) < abs(price_buy_change)){
            #PS_change = PS_change-(abs(abs(price_buy_change) - abs(price_sell_change)))*(mean(hour_buy_curve$hour_buy_q))
            #}
            #welfare_change = PS_change+CS_change
          }
          #correction for imprecise sub_curve selection
          CS_change = CS_change - (price_change+price_buy_change)*(mean(sub_buy_curve$hour_buy_q))
          PS_change = PS_change + (price_change-price_sell_change)*(mean(sub_sell_curve$hour_sell_q))
          welfare_change = PS_change+CS_change
          
          
          intersection = cbind(intersection, price_buy_change,CS_change,price_sell_change,PS_change,welfare_change)
          
          names(intersection) = c("time","original_price","original_quantity","german_price", "new_price", "new_quantity", "ger_nor_org_diff", "price_change", "perc_diff", "price_buy_change","CS_change","price_sell_change","PS_change","welfare_change")
          #renames the columns in intersection result
          
          intersections = rbind(intersections, intersection)
          #writes the result into a dataframe
          
          #assign(paste(hour,"sub_buy_curve",sep=""),sub_buy_curve) 
          #assign(paste(hour,"sub_sell_curve",sep=""),sub_sell_curve)
          
          #Plot generator: !!! generates shifted curve before any adjustments (parameter, stop change etc.) 
          #ggplot()+
          #geom_line(data=shift_hour_buy_curve, aes(x=shift_hour_buy_curve$hour_buy_q, y=shift_hour_buy_curve$hour_buy_price), color='green',size=0.35)+
          #geom_line(data=shift_hour_sell_curve, aes(x=shift_hour_sell_curve$hour_sell_q, y=shift_hour_sell_curve$hour_sell_price), color='green',size=0.35)+
          #geom_line(data=hour_buy_curve, aes(x=hour_buy_curve$hour_buy_q, y=hour_buy_curve$hour_buy_price), color='black',size=0.5) + 
          #geom_line(data=hour_sell_curve, aes(x=hour_sell_curve$hour_sell_q, y=hour_sell_curve$hour_sell_price), color='black',size=0.5)+
          #xlab("Quantity MWh") + ylab("Price EUR") + ggtitle(paste(capacity_MW,"_",robustness,"_",substring(data_copy[1,hour*2],7,10),substring(data_copy[1,hour*2],3,6),substring(data_copy[1,hour*2],1,2),substring(data_copy[1,hour*2],11,19),sep = ""))+ theme_bw() + theme(panel.border = element_blank(), panel.grid.major = element_blank(), panel.grid.minor = element_blank(), axis.line = element_line(colour = "black"))+
          #xlim(20000,70000)
          
          #ggsave(paste(capacity_MW,"_",robustness,"_",substring(data_copy[1,hour*2],7,10),substring(data_copy[1,hour*2],3,6),substring(data_copy[1,hour*2],1,2),substring(data_copy[1,hour*2],11,19),".pdf",sep = ""), device = "pdf", path = paste("~/Desktop/electricity/graphs/",year,"/",sep = ""))
        }
      }
      
      write.xlsx(intersections, paste("intersections_",country,year,"_",capacity_MW,"_",robustness,".xlsx",sep="") )
      
      #POSTANALYSIS:  (for the whole year)
      
      #calculate the change of average price (averaged by new volume):
      avg_calc=data.frame(intersections[,6]*intersections[,8])
      avg_price_change = sum(avg_calc)/sum(intersections[,6])
      
      #calculate the change of the difference between foreign and Norwegian price (averaged by new volume):
      avg_calc=data.frame(intersections[,6]*intersections[,7])
      avg_ger_nor_price_diff = sum(avg_calc)/sum(intersections[,6])
      
      #calculate absolute value of the change of the difference between foreign and Norwegian price (averaged by new volume):
      avg_calc=data.frame(intersections[,6]*abs(intersections[,7]))
      avg_ABS_ger_nor_price_diff = sum(avg_calc)/sum(intersections[,6])
      
      #calculate the average Original Norwegian price (averaged by new volume):
      avg_calc=data.frame(intersections[,6]*intersections[,2])
      avg_org_nor_price = sum(avg_calc)/sum(intersections[,6])
      
      #calculate the average changed Norwegian price (averaged by new volume):
      avg_calc=data.frame(intersections[,6]*intersections[,5])
      avg_new_nor_price = sum(avg_calc)/sum(intersections[,6])
      
      #calculate the average foreign price (averaged by new Norwegian volume):
      avg_calc=data.frame(intersections[,6]*intersections[,4])
      avg_ger_price = sum(avg_calc)/sum(intersections[,6])
      
      #calculate the average perc_diff (averaged by new volume):
      avg_calc=data.frame(intersections[,6]*intersections[,9])
      avg_perc_diff_weighted = sum(avg_calc)/sum(intersections[,6])
      
      #calculate average perc_diff (averaged by hours):
      avg_percent_diff_nonweighted = mean(intersections[,9])
      
      #calculate sum price_change:
      sum_price_sell_change = sum(intersections$price_sell_change)
      
      #calculate sum price_buy_change:
      sum_price_buy_change = sum(intersections$price_buy_change)
      
      #calculate sum PS_change:
      sum_PS_change = sum(intersections$PS_change)
      
      #calculate sum CS_change
      sum_CS_change = sum(intersections$CS_change)
      
      #calculate sum welfare_change
      sum_welfare_change = sum(intersections$welfare_change)
      
      result = data.frame(year, capacity_MW, robustness)
      #crates a dataframe for result of this itteration 
      result = cbind(result,avg_price_change,avg_org_nor_price,avg_ger_price,avg_new_nor_price,avg_ger_nor_price_diff,avg_ABS_ger_nor_price_diff,avg_perc_diff_weighted,avg_percent_diff_nonweighted,sum_price_sell_change,sum_price_buy_change,sum_PS_change,sum_CS_change,sum_welfare_change,country)
      #puts the results into dataframe
      names(result) = c("year","capacity_MW","robustness","avg_price_change","avg_org_nor_price","avg_ger_price","avg_new_nor_price","avg_ger_nor_price_diff","avg_ABS_ger_nor_price_diff","avg_perc_diff_weighted","avg_percent_diff_nonweighted","sum_price_sell_change","sum_price_buy_change","sum_PS_change","sum_CS_change","sum_welfare_change","country")
      #renames the columns in result
      results = rbind(results,result)
      #binds the result into results dataframe
    }
  }
}
write.xlsx(results, "results.xlsx" )
#saves the results as an excel file 



