#Stages Timesheet info
#make sure Timesheet and Expenses Reports are sorted correctly
require 'rubygems'
require 'roo'
require 'write_xlsx'
require_relative 'Methods.rb'
require 'date'



def stageTimesheet

  eraseFile(getFile('st_ts'), "Timesheet")

  #open files, set defult sheet

  ts = Roo::Excelx.new(getFile('ts'))
  ms = Roo::Excelx.new(getFile('ms'))
  ms.default_sheet = ms.sheets[0]
  ts.default_sheet = ts.sheets[0]


  stage = WriteXLSX.new(getFile('st_ts'))
  timesheet = stage.add_worksheet('Timesheet')




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


  #set tsData array
  tsData = Array.new(28){Array.new()}
  #fill tsData array
  ts.each do |row|
    if row[10].instance_of?(String)  #check for cost enter code
      if row[10].length == 4 #make sure cost center code is 4 digits long
        tsData[0].push(row[9])#worker name
        tsData[1].push()#VMSID empty till master roster match
        tsData[2].push(StateSplit(row[0]))#state
        tsData[3].push(row[10])#cost center
        tsData[4].push(row[11])#week end date
        tsData[5].push(row[3])#bill rate
        tsData[6].push(row[4])#bill overtime rate
        tsData[7].push(row[5])#bill doubletime rate
        tsData[8].push(row[1])#time sheet status
        tsData[9].push(row[6])#billable ST/Hr
        tsData[10].push(row[7])#billable OT/Hr
        tsData[11].push(row[8])#billable DT/Hr
        tsData[12].push(0)#mpci_claims
        tsData[13].push(0)#mpci_train
        tsData[14].push(0)#nau
        tsData[15].push(0)#hail
        tsData[16].push(0)#hail_train
        tsData[17].push(0)#underwriting
        tsData[18].push(0)#retirement
        tsData[19].push(0)#esl
        tsData[20].push(0)#na
        tsData[21].push()#total hours
        tsData[22].push(0)#total ST/Hr cost
        tsData[23].push(0)#total OT/Hr cost
        tsData[24].push(0)#total DT/Hr cost
        tsData[25].push(0)#total hours cost
        tsData[26].push(row[2])#task name
        tsData[27].push(true)#flag
      end
    end
  end

  #get VMSIDs
  tsData[0].each_index do |i|
    ms_name.each_index do |j|
      if tsData[0][i] == ms_name[j]
        tsData[1][i] = ms_vmsid[j]
      end
    end
  end

  task_name = ["(50030) MPCI Claims","(50034) MPCI Training","(50032) NAU Compliance","(50031) Hail/Named Peril","(50035) Hail Training","(50033) Underwriting Inspection","(37379) Mobius Retirement","(37417) ESL Net Migration"]


   #["(50030) MPCI Claims","(00000) No project-allocate on time","(CS412) Topeka (Fielding Rd)","(22757) ENH Claims","(50034) MPCI Training","(50032) NAU Compliance","(50031) Hail/Named Peril","(TRAIN) Training expenses","(50054) Crop Claims Meet & Training","(50033) Underwriting Inspection","(CS409) Eau Claire","(37379) Mobius Retirement","(37417) ESL Net Migration" ]



  #erase draft Hours
  tsData[0].each_index do |i|
    if tsData[8][i] == "Draft"
      tsData[9][i] = 0
      tsData[10][i] = 0
      tsData[11][i] = 0
    end
  end

  #separate tsData hours by task name
  tsData[0].each_index do |i|
    hours = tsData[9][i] + tsData[10][i] + tsData[11][i]

    tn = task_name.index{|s| s.include?(tsData[26][i].to_s)}
    if tn == 0 then tsData[12][i] = hours end
    if tn == 1 then tsData[13][i] = hours end
    if tn == 2 then tsData[14][i] = hours end
    if tn == 3 then tsData[15][i] = hours end
    if tn == 4 then tsData[16][i] = hours end
    if tn == 5 then tsData[17][i] = hours end
    if tn == 6 then tsData[18][i] = hours end
    if tn == 7 then tsData[19][i] = hours end
    if tn == nil then tsData[20][i] = hours end
    tsData[21][i] = hours
    hours = 0
  end
  #combine hours by week end date
  tsData[0].each_index do |i|
    if tsData[0][i-1]
      if tsData[0][i] == tsData[0][i-1] && tsData[4][i] == tsData[4][i-1]
        tn2 = task_name.index{|s| s.include?(tsData[26][i].to_s)}
        if tn2 == 0
          tsData[12][i] = tsData[12][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == 1
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == 2
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == 3
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == 4
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == 5
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == 6
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == 7
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
          tsData[20][i] = tsData[20][i-1]
        end
        if tn2 == nil
          tsData[12][i] = tsData[12][i-1]
          tsData[13][i] = tsData[13][i-1]
          tsData[14][i] = tsData[14][i-1]
          tsData[15][i] = tsData[15][i-1]
          tsData[16][i] = tsData[16][i-1]
          tsData[17][i] = tsData[17][i-1]
          tsData[18][i] = tsData[18][i-1]
          tsData[19][i] = tsData[19][i-1]
          tsData[20][i] = tsData[20][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
        end
        tsData[21][i] = tsData[21][i-1] + tsData[9][i] + tsData[10][i] + tsData[11][i]
        tsData[27][i-1] = false
      end
    end
  end



  #Header
  #header = [ "Worker", "VMSID", "State", "Cost Center", "Crop Year", "Expense Ending Period" ,"Hourly Bill Rate $", "Hourly Overtime Bill Rate $", "Time Sheet Status", "MPCI Claims", "No project-allocate on time", "Topeka (Fielding Rd)", "ENH Claims", "MPCI Training", "NAU Compliance", "Hail/Named Peril", "Training Expenses", "Crop Claims Meet & Training", "Underwriting Inspection",
  #  "Eau Claire","Mobius Retirement","ESL Net Migration", "N/A","Total Regular Hours", "Total Overtime Hours","Total Hours","Total Regular Hours Cost $", "Total Overtime Hours Cost $", "Total Hours Cost $"]


  #write header
  #header.each_index do |i|
#    timesheet.write(0, i, header[i])
#  end

  #setup hours cost
  tsData[0].each_index do |i|
    tsData[22][i] = tsData[9][i] * tsData[5][i]#ST/Hr cost = billable ST/Hr * ST/Hr bill rate
    tsData[23][i] = tsData[10][i] * tsData[6][i]#DT/Hr cost = billable DT/Hr * DT/Hr bill rate
    tsData[24][i] = tsData[11][i] * tsData[7][i]#OT/Hr cost = billable OT/Hr * OT/Hr bill rate
    tsData[25][i] = tsData[22][i] + tsData[23][i]  + tsData[24][i]#Total Hours cost = ST/Hr cost + OT/Hr cost + DT/Hr cost
  end



  count = 0
  tsData[0].each_index do |i|
    if tsData[27][i] == true && getCropYear(tsData[4][i]).to_i > 2017 && tsData[4][i] < getCurrentDate()
      timesheet.write(count, 0, tsData[0][i])#name
      timesheet.write(count, 1, tsData[1][i])#vmsid
      timesheet.write(count, 2, tsData[2][i])#state
      timesheet.write(count, 3, tsData[3][i])#cost center
      timesheet.write(count, 4, getCropYear(tsData[4][i]))#crop year
      timesheet.write(count, 5, tsData[4][i])#expense ending period
      timesheet.write(count, 6, tsData[5][i])#bill rate
      timesheet.write(count, 7, tsData[6][i])#ot bill rate
      timesheet.write(count, 8, tsData[7][i])#dt bill rate
      timesheet.write(count, 9, tsData[8][i])#time sheet status
      timesheet.write(count, 10, tsData[12][i])#mpci_claims
      timesheet.write(count, 11, tsData[13][i])#mpci_train
      timesheet.write(count, 12, tsData[14][i])#nau
      timesheet.write(count, 13, tsData[15][i])#hail
      timesheet.write(count, 14, tsData[16][i])#hail_train
      timesheet.write(count, 15, tsData[17][i])#underwriting
      timesheet.write(count, 16, tsData[18][i])#retirement
      timesheet.write(count, 17, tsData[19][i])#ESL Net Migration
      timesheet.write(count, 18, tsData[20][i])#na
      timesheet.write(count, 19, tsData[9][i])#total ST hours
      timesheet.write(count, 20, tsData[10][i])#total OT hours
      timesheet.write(count, 21, tsData[11][i])#total DT hours
      timesheet.write(count, 22, tsData[21][i])#total hours
      timesheet.write(count, 23, tsData[22][i])#total reg hours cost
      timesheet.write(count, 24, tsData[23][i])#total ot hours cost
      timesheet.write(count, 25, tsData[24][i])#total DT hours cost
      timesheet.write(count, 26, tsData[25][i])#total hours cost
      count += 1
    end
  end




  stage.close



end
