#Rollup

require 'rubygems'
require 'roo'
require 'write_xlsx'
require_relative 'Methods.rb'
require 'date'


def createRollup()

  eraseFile(getFile('rl'),'Invoice')




  #open files, set defult sheet
  st_ts = Roo::Excelx.new(getFile('st_ts'))
  st_ex = Roo::Excelx.new(getFile('st_ex'))
  st_ts.default_sheet = st_ts.sheets[0]
  st_ex.default_sheet = st_ex.sheets[0]

  rollup = WriteXLSX.new(getFile('rl'))
  invoice = rollup.add_worksheet('Invoice')

  #fill tsData array
  tsData = Array.new(28){Array.new()}

  st_ts.each do |row|
      tsData[0].push(row[0])#worker name
      tsData[1].push(row[1])#VMSID
      tsData[2].push(row[2])#state
      tsData[3].push(row[3])#cost center
      tsData[4].push(row[4])#crop year
      tsData[5].push(row[5])#week end date
      tsData[6].push(row[6])#ST bill rate
      tsData[7].push(row[7])#OT bill rate
      tsData[8].push(row[8])#DT bill rate
      tsData[9].push(row[9])#time sheet status
      tsData[10].push(row[10])#mpci_claims
      tsData[11].push(row[11])#mpci_train
      tsData[12].push(row[12])#nau
      tsData[13].push(row[13])#hail
      tsData[14].push(row[14])#hail_train
      tsData[15].push(row[15])#underwriting
      tsData[16].push(row[16])#retirement
      tsData[17].push(row[17])#esl
      tsData[18].push(row[18])#na
      tsData[19].push(row[19])#total reg hours
      tsData[20].push(row[20])#total ot hours
      tsData[21].push(row[21])#total dt hours
      tsData[22].push(row[22])#total hours
      tsData[23].push(row[23])#total ST hours cost
      tsData[24].push(row[24])#total OT hours cost
      tsData[25].push(row[25])#total DT hours cost
      tsData[26].push(row[26])#total hours cost
      tsData[27].push(true)#flag
  end


  #fill exData array
  exData = Array.new(22){Array.new()}
  st_ex.each do |row|
      exData[0].push(row[0])#worker name
      exData[1].push(row[1])#VMSID
      exData[2].push(row[2])#state
      exData[3].push(row[3])#cost center
      exData[4].push(row[4])#crop year
      exData[5].push(row[5])#week end date
      exData[6].push(row[6])#milage
      exData[7].push(row[7])#cell Phone
      exData[8].push(row[8])#internet
      exData[9].push(row[9])#hotel
      exData[10].push(row[10])#tips & gratuities
      exData[11].push(row[11])#office Supplies
      exData[12].push(row[12])#training
      exData[13].push(row[13])#Airfare
      exData[14].push(row[14])#Airfare - Baggage fees
      exData[15].push(row[15])#Taxi/Bus Fare
      exData[16].push(row[16])#Parking/Tolls
      exData[17].push(row[17])#Unspecified transaction type
      exData[18].push(row[18])#Meals (Breakfast, Lunch, Dinner)
      exData[19].push(row[19])#N/A
      exData[20].push(row[20])#Total
      exData[21].push(true)#flag
  end


  header = ["Worker", "VMSID", "State", "Cost Center", "Crop Year", "Expense Ending Period" ,"Hourly Bill Rate $", "Hourly Overtime Bill Rate $", "Hourly Doubletime Bill Rate $", "Time Sheet Status", "MPCI Claims (Hours)", "MPCI Training (Hours)", "NAU Compliance (Hours)", "Hail/Named Peril (Hours)","Hail Training (Hours)", "Underwriting Inspection (Hours)", "Mobius Retirement (Hours)", "ESL Net Migration (Hours)", "N/A (Hours)", "Total Regular Hours", "Total Overtime Hours", "Total Doubletime Hours", "Total Hours", "Personal Mileage $", "Cell Phone $", "Internet $", "Hotel $", "Tips & Gratuities $", "Office Supplies $",
    "Training $", "Airfare $", "Airfare - Baggage Fees $", "Taxi/Bus Fare $", "Parking/Tolls $", "Unspecified transaction type $", "Meals (Breakfast, Lunch, Dinner) $", "N/A$", "Total Expenses $", "Total Regular Hours Cost $", "Total Overtime Hours Cost $","Total Doubletime Hours Cost $", "Total Hours Cost $", "Total Cost $"]


  #write header
  header.each_index do |i|
    invoice.write(0, i, header[i])
  end


  count = 1


  #joint output
  tsData[0].each_index do |i|
    exData[0].each_index do |j|
      if tsData[0][i] == exData[0][j] #check names match
        if tsData[5][i] == exData[5][j] #check week ends match
          invoice.write(count, 0, tsData[0][i])#name
          invoice.write(count, 1, tsData[1][i])#vmsid
          invoice.write(count, 2, tsData[2][i])#state
          invoice.write(count, 3, tsData[3][i])#cost center
          invoice.write(count, 4, tsData[4][i])#crop year
          invoice.write(count, 5, tsData[5][i])#expense ending period
          invoice.write(count, 6, tsData[6][i])#bill rate
          invoice.write(count, 7, tsData[7][i])#ot bill rate
          invoice.write(count, 8, tsData[8][i])#dt bill rate
          invoice.write(count, 9, tsData[9][i])#time sheet status
          invoice.write(count, 10, tsData[10][i])#mpci_claims
          invoice.write(count, 11, tsData[11][i])#mpci_train
          invoice.write(count, 12, tsData[12][i])#nau
          invoice.write(count, 13, tsData[13][i])#hail
          invoice.write(count, 14, tsData[14][i])#hail_train
          invoice.write(count, 15, tsData[15][i])#underwriting
          invoice.write(count, 16, tsData[16][i])#retirement
          invoice.write(count, 17, tsData[17][i])#esl
          invoice.write(count, 18, tsData[18][i])#na
          invoice.write(count, 19, tsData[19][i])#total reg hours
          invoice.write(count, 20, tsData[20][i])#total ot hours
          invoice.write(count, 21, tsData[21][i])#total dt hours
          invoice.write(count, 22, tsData[22][i])#total Hours
          invoice.write(count, 23, exData[6][j])#mileage
          invoice.write(count, 24, exData[7][j])#cell phone
          invoice.write(count, 25, exData[8][j])#internet
          invoice.write(count, 26, exData[9][j])#Hotel
          invoice.write(count, 27, exData[10][j])#Tips & Gratuities
          invoice.write(count, 28, exData[11][j])#Office Supplies
          invoice.write(count, 29, exData[12][j])#Training
          invoice.write(count, 30, exData[13][j])#Airfare
          invoice.write(count, 31, exData[14][j])#Airfare - Baggage Fees
          invoice.write(count, 32, exData[15][j])#Taxi/Bus Fare
          invoice.write(count, 33, exData[16][j])#Parking/Tolls
          invoice.write(count, 34, exData[17][j])#Unspecified transaction type
          invoice.write(count, 35, exData[18][j])#Meals (Breakfast, Lunch, Dinner)
          invoice.write(count, 36, exData[19][j])#n/a
          invoice.write(count, 37, exData[20][j])#total expenses
          invoice.write(count, 38, tsData[23][i])#total reg hours cost
          invoice.write(count, 39, tsData[24][i])#total ot hours cost
          invoice.write(count, 40, tsData[25][i])#total dt hours cost
          invoice.write(count, 41, tsData[26][i])#total hours cost
          total_cost = tsData[26][i].to_f + exData[20][j].to_f
          invoice.write(count, 42, total_cost)#total cost
          tsData[27][i] = false
          exData[21][j] = false
          count += 1
        end
      end
    end
  end



  #timesheet only
  tsData[0].each_index do |i|
    if tsData[24][i] == true && tsData[20][i] > 0
      invoice.write(count, 0, tsData[0][i])#name
      invoice.write(count, 1, tsData[1][i])#vmsid
      invoice.write(count, 2, tsData[2][i])#state
      invoice.write(count, 3, tsData[3][i])#cost center
      invoice.write(count, 4, tsData[4][i])#crop year
      invoice.write(count, 5, tsData[5][i])#expense ending period
      invoice.write(count, 6, tsData[6][i])#bill rate
      invoice.write(count, 7, tsData[7][i])#ot bill rate
      invoice.write(count, 8, tsData[8][i])#dt bill rate
      invoice.write(count, 9, tsData[9][i])#time sheet status
      invoice.write(count, 10, tsData[10][i])#mpci_claims
      invoice.write(count, 11, tsData[11][i])#mpci_train
      invoice.write(count, 12, tsData[12][i])#nau
      invoice.write(count, 13, tsData[13][i])#hail
      invoice.write(count, 14, tsData[14][i])#hail_train
      invoice.write(count, 15, tsData[15][i])#underwriting
      invoice.write(count, 16, tsData[16][i])#retirement
      invoice.write(count, 17, tsData[17][i])#esl
      invoice.write(count, 18, tsData[18][i])#na
      invoice.write(count, 19, tsData[19][i])#total reg hours
      invoice.write(count, 20, tsData[20][i])#total ot hours
      invoice.write(count, 21, tsData[21][i])#total dt hours
      invoice.write(count, 22, tsData[22][i])#total Hours
      invoice.write(count, 23, 0)#mileage
      invoice.write(count, 24, 0)#cell phone
      invoice.write(count, 25, 0)#internet
      invoice.write(count, 26, 0)#Hotel
      invoice.write(count, 27, 0)#Tips & Gratuities
      invoice.write(count, 28, 0)#Office Supplies
      invoice.write(count, 29, 0)#Training
      invoice.write(count, 30, 0)#Airfare
      invoice.write(count, 31, 0)#Airfare - Baggage Fees
      invoice.write(count, 32, 0)#Taxi/Bus Fare
      invoice.write(count, 33, 0)#Parking/Tolls
      invoice.write(count, 34, 0)#Unspecified transaction type
      invoice.write(count, 35, 0)#Meals (Breakfast, Lunch, Dinner)
      invoice.write(count, 36, 0)#n/a
      invoice.write(count, 37, 0)#total expenses
      invoice.write(count, 38, tsData[23][i])#total reg hours cost
      invoice.write(count, 39, tsData[24][i])#total ot hours cost
      invoice.write(count, 40, tsData[25][i])#total dt hours cost
      invoice.write(count, 41, tsData[26][i])#total hours cost
      invoice.write(count, 42, tsData[26][i])#total cost
      count += 1
    end
  end



  #expenses only
  exData[0].each_index do |i|
    if exData[21][i] == true && exData[20][i] > 0
      invoice.write(count, 0, exData[0][i])#name
      invoice.write(count, 1, exData[1][i])#vmsid
      invoice.write(count, 2, exData[2][i])#state
      invoice.write(count, 3, exData[3][i])#cost center
      invoice.write(count, 4, exData[4][i])#crop year
      invoice.write(count, 5, exData[5][i])#expense ending period
      invoice.write(count, 6, 0)#bill rate
      invoice.write(count, 7, 0)#ot bill rate
      invoice.write(count, 8, 0)#dt bill rate
      invoice.write(count, 9, "Expense with no Time Sheet")#time sheet status
      invoice.write(count, 11, 0)#mpci_train
      invoice.write(count, 12, 0)#nau
      invoice.write(count, 13, 0)#hail
      invoice.write(count, 14, 0)#hail_train
      invoice.write(count, 15, 0)#underwriting
      invoice.write(count, 16, 0)#retirement
      invoice.write(count, 17, 0)#esl
      invoice.write(count, 18, 0)#na
      invoice.write(count, 19, 0)#total reg hours
      invoice.write(count, 20, 0)#total ot hours
      invoice.write(count, 21, 0)#total dt hours
      invoice.write(count, 22, 0)#total Hours
      invoice.write(count, 23, exData[6][i])#mileage
      invoice.write(count, 24, exData[7][i])#cell phone
      invoice.write(count, 25, exData[8][i])#internet
      invoice.write(count, 26, exData[9][i])#Hotel
      invoice.write(count, 27, exData[10][i])#Tips & Gratuities
      invoice.write(count, 28, exData[11][i])#Office Supplies
      invoice.write(count, 29, exData[12][i])#Training
      invoice.write(count, 30, exData[13][i])#Airfare
      invoice.write(count, 31, exData[14][i])#Airfare - Baggage Fees
      invoice.write(count, 32, exData[15][i])#Taxi/Bus Fare
      invoice.write(count, 33, exData[16][i])#Parking/Tolls
      invoice.write(count, 34, exData[17][i])#Unspecified transaction type
      invoice.write(count, 35, exData[18][i])#Meals (Breakfast, Lunch, Dinner)
      invoice.write(count, 36, exData[19][i])#n/a
      invoice.write(count, 37, exData[20][i])#total expenses
      invoice.write(count, 38, 0)#total reg hours cost
      invoice.write(count, 39, 0)#total ot hours cost
      invoice.write(count, 40, 0)#total dt hours cost
      invoice.write(count, 41, 0)#total hours cost
      invoice.write(count, 42, exData[20][i])#total cost
      count += 1
    end
  end



  rollup.close

end
