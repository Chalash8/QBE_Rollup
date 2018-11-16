require_relative 'Program Files/Timesheet.rb'
require_relative 'Program Files/Expenses.rb'
require_relative 'Program Files/Rollup.rb'
require_relative 'Program Files/Tabs.rb'




puts "Make sure Master Roster is updated! "
print "Make sure Timesheet and Expense Reports are already sorted! "
gets


stageTimesheet()
puts "Timesheet.rb complete"
stageExpenses()
puts "Expenses.rb complete"
createRollup()
puts "Rollup.rb complete"
