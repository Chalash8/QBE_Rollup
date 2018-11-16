#Stages Expense info
#make sure expenses and Expenses Reporex are sorted correctly
require 'rubygems'
require 'roo'
require 'write_xlsx'
require_relative 'Methods.rb'
require 'date'



def stageExpenses

  eraseFile(getFile('st_ex'), "Expenses")

  #open files, set defult sheet
  ex = Roo::Excelx.new(getFile('ex'))
  ms = Roo::Excelx.new(getFile('ms'))
  ex.default_sheet = ex.sheets[0]
  ms.default_sheet = ms.sheets[0]


  stage = WriteXLSX.new(getFile('st_ex'))
  expenses = stage.add_worksheet('Expenses')


  #set ms arrays
  ms_name = []
  ms_vmsid = []
  #fill ms arrays
  ms.each do |row|
    if row[1] == 'QBE-FG'
      ms_name.push(row[0])
      ms_vmsid.push(row[2])
    end
  end




  #set exData array
  exData = Array.new(24){Array.new()}
  #fill exData array
  ex.each do |row|
    if row[3].instance_of?(String)  #check for cost enter code
      if row[3].length == 4 #make sure cost center code is 4 digits long
        exData[0].push(row[0])#worker name
        exData[1].push(getWeekEnding(row[12]))#week end date
        exData[2].push(row[1])#billable expenses
        exData[3].push(row[9])#expense code
        exData[4].push(0)#milage
        exData[5].push(0)#cell Phone
        exData[6].push(0)#internet
        exData[7].push(0)#hotel
        exData[8].push(0)#tips & gratuities
        exData[9].push(0)#office Supplies
        exData[10].push(0)#training
        exData[11].push(0)#Airfare
        exData[12].push(0)#Airfare - Baggage fees
        exData[13].push(0)#Taxi/Bus Fare
        exData[14].push(0)#Parking/Tolls
        exData[15].push(0)#Unspecified transaction type
        exData[16].push(0)#Meals (Breakfast, Lunch, Dinner)
        exData[17].push(0)#N/A
        exData[18].push(0)#Total Expenses
        exData[19].push(true)#flag
        exData[20].push(StateSplit(row[10]))#state
        exData[21].push(row[3])#cost center
        exData[22].push(getCropYear(row[12]))#crop year
        exData[23].push()#VMSID
      end
    end
  end


  #get VMSIDs
  exData[0].each_index do |i|
    ms_name.each_index do |j|
      if exData[0][i] == ms_name[j]
        exData[23][i] = ms_vmsid[j]
      end
    end
  end



  expense_code = ["TRNS014 - Personal Mileage","COMM003 - Cell Phone","COMM001 - Internet","LODG001 - Hotel","MISC001 - Tips & Gratuities","MISC006 - Office Supplies","MISC013 - Training","TRNS006 - Airfare","TRNS010 - Airfare - Baggage Fees","TRNS012 - Taxi/Bus Fare","TRNS015 - Parking/Tolls","UNSPC - Unspecified transaction type","MEAL001 - Meals - Breakfast","MEAL002 - Meals - Lunch","MEAL003 - Meals - Dinner"]


  #separate exData expenses by expense code
  exData[0].each_index do |i|
    ec = expense_code.index{|s| s.include?(exData[3][i].to_s)}
    if ec == 0 then exData[4][i] = exData[2][i] end
    if ec == 1 then exData[5][i] = exData[2][i] end
    if ec == 2 then exData[6][i] = exData[2][i] end
    if ec == 3 then exData[7][i] = exData[2][i] end
    if ec == 4 then exData[8][i] = exData[2][i] end
    if ec == 5 then exData[9][i] = exData[2][i] end
    if ec == 6 then exData[10][i] = exData[2][i] end
    if ec == 7 then exData[11][i] = exData[2][i] end
    if ec == 8 then exData[12][i] = exData[2][i] end
    if ec == 9 then exData[13][i] = exData[2][i] end
    if ec == 10 then exData[14][i] = exData[2][i] end
    if ec == 11 then exData[15][i] = exData[2][i] end
    if ec == 12 then exData[15][i] = exData[2][i] end
    if ec == 13 then exData[15][i] = exData[2][i] end
    if ec == 14 then exData[16][i] = exData[2][i] end
    if ec == nil then exData[17][i] = exData[2][i] end
    exData[18][i] = exData[2][i]
  end



  #combine expenses by week end date
  exData[0].each_index do |i|
    if exData[0][i-1]
      if exData[0][i] == exData[0][i-1] && exData[1][i] == exData[1][i-1]
        ec = expense_code.index{|s| s.include?(exData[3][i].to_s)}
        if ec == 0
          exData[4][i] = exData[4][i-1] + exData[2][i]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 1
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1] + exData[2][i]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 2
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1] + exData[2][i]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 3
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1] + exData[2][i]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 4
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1] + exData[2][i]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 5
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1] + exData[2][i]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 6
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1] + exData[2][i]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 7
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1] + exData[2][i]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 8
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1] + exData[2][i]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 9
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1] + exData[2][i]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 10
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1] + exData[2][i]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 11
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1] + exData[2][i]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 12
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1] + exData[2][i]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 13
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1] + exData[2][i]
          exData[17][i] = exData[17][i-1]
        end
        if ec == 14
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1] + exData[2][i]
          exData[17][i] = exData[17][i-1]
        end
        if ec == nil
          exData[4][i] = exData[4][i-1]
          exData[5][i] = exData[5][i-1]
          exData[6][i] = exData[6][i-1]
          exData[7][i] = exData[7][i-1]
          exData[8][i] = exData[8][i-1]
          exData[9][i] = exData[9][i-1]
          exData[10][i] = exData[10][i-1]
          exData[11][i] = exData[11][i-1]
          exData[12][i] = exData[12][i-1]
          exData[13][i] = exData[13][i-1]
          exData[14][i] = exData[14][i-1]
          exData[15][i] = exData[15][i-1]
          exData[16][i] = exData[16][i-1]
          exData[17][i] = exData[17][i-1] + exData[2][i]
        end
        exData[18][i] = exData[18][i-1] + exData[2][i]
        exData[19][i-1] = false
      end
    end
  end

  #header = ["Worker","Expense Ending Period","Personal Mileage","Cell Phone", "Internet", "Hotel","Tips & Gratuities","Office Supplies","Training","Airfare","Airfare - Baggage Fees","Taxi/Bus Fare","Parking/Tolls","Unspecified transaction type","Meals (Breakfast, Lunch, Dinner)","N/A","Total Expenses"]

  #write header
  #header.each_index do |i|
    #expenses.write(0, i, header[i])
  #end


  count = 0
  exData[0].each_index do |i|
    if exData[19][i] == true && exData[22][i].to_i > 2017 && exData[1][i] < getCurrentDate()
      expenses.write(count, 0, exData[0][i])#name
      expenses.write(count, 1, exData[23][i])#VMSID
      expenses.write(count, 2, exData[20][i])#state
      expenses.write(count, 3, exData[21][i])#cost center
      expenses.write(count, 4, exData[22][i])#crop year
      expenses.write(count, 5, exData[1][i])#expense ending period
      expenses.write(count, 6, exData[4][i])#mileage
      expenses.write(count, 7, exData[5][i])#cell phone
      expenses.write(count, 8, exData[6][i])#internet
      expenses.write(count, 9, exData[7][i])#Hotel
      expenses.write(count, 10, exData[8][i])#Tips & Gratuities
      expenses.write(count, 11, exData[9][i])#Office Supplies
      expenses.write(count, 12, exData[10][i])#Training
      expenses.write(count, 13, exData[11][i])#Airfare
      expenses.write(count, 14, exData[12][i])#Airfare - Baggage Fees
      expenses.write(count, 15, exData[13][i])#Taxi/Bus Fare
      expenses.write(count, 16, exData[14][i])#Parking/Tolls
      expenses.write(count, 17, exData[15][i])#Unspecified transaction type
      expenses.write(count, 18, exData[16][i])#Meals (Breakfast, Lunch, Dinner)
      expenses.write(count, 19, exData[17][i])#n/a
      expenses.write(count, 20, exData[18][i])#total
      count += 1
    end
  end


  stage.close

  end
