require 'date'
require 'roo'
require 'write_xlsx'


def getCurrentDate()
  return DateTime.now
end


def getFile(file)
  return '/Users/chadhalash/GoogleDriveChad/QBE/Pull Files/Time.xlsx' if file == 'ts'
  return '/Users/chadhalash/GoogleDriveChad/QBE/Pull Files/Expenses.xlsx' if file == 'ex'
  return '/Users/chadhalash/GoogleDriveChad/QBE/Pull Files/Master Roster.xlsx' if file == 'ms'
  return '/Users/chadhalash/GoogleDriveChad/QBE/Static Stage Timesheet.xlsx' if file == 'st_ts'
  return '/Users/chadhalash/GoogleDriveChad/QBE/Static Stage Expenses.xlsx' if file == 'st_ex'
  return '/Users/chadhalash/GoogleDriveChad/QBE/Static Rollup.xlsx' if file == 'rl'
end

def eraseFile(file, sheet)
  fileRoo = Roo::Excelx.new(file)
  fileRoo.default_sheet = sheet

  fileWrite = WriteXLSX.new(file)
  fileSheet = fileWrite.add_worksheet(sheet)

  header = Array.new(44)

  if fileRoo.first_row
    fileRoo.each_with_index do |i|
      header.each_index do |j|
        fileSheet.write(i, j, "")
      end
    end
  end
  fileSheet.write(0, 0, sheet)
  fileWrite.close

  out = sheet.to_s + " Cleared"
  puts out

end

def StateSplit(state)
  state = state.to_s
  if state == 'USA - VT - Remote'
    return state[6,2]
  else
    return state[0,2]
  end
end

def getCropYear(date)
  return date.to_s[0,4]
end

def getHourlyRate(rate)
  rate = rate.to_f / 1.2
  return rate
end

def getOvertimeRate(rate)
  rate = getHourlyRate(rate) * 1.5
  return rate
end

def getWeekEnding(date)
  if date.instance_of?(Date)
    date = date + ((0 - date.wday) % 7)
  end
  return date
end
