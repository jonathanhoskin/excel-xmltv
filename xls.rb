require 'rubygems'
require 'spreadsheet'
require 'roo'
require 'parsedate'

# 'spreadsheet' gem test
# book = Spreadsheet.open 'triangle.xls'
# sheet = book.worksheet 0
# 
# sheet.each do |row|
#   puts row.join(',')
#   format = row.format 3
#   # puts format.to_a
# end

# 'roo'
daterow = 2
startrow = 3
lastrow = 52
startcolumn = "b"
firstemptycolumn = "i"
timecolumn = "a"
file = "triangle.xls"
column = startcolumn
info = ['[PGR]', '[AO]']


s = Excel.new(file)
s.default_sheet = s.sheets.first

# Convert seconds into hours and minutes
def convert_time(secs)
  mins = secs / 60
  hours = ((mins % (60 *24)) / 60).floor
  hours = "%02d" % hours
  minutes = mins % 60
  minutes = "%02d" % minutes
  [hours, minutes].join
end

def check_info(info, data)
  info.each do |i|
#    puts "i #{i}"
    data.split.each do |d|
#      puts "d #{d}"
      return data if d == i
    end
  end
  return
end

# # This is the date
# until column == "i"
#   cell = s.cell(column,daterow)
#   day = ParseDate.parsedate cell, true
#   puts "#{day[2]},#{day[1]},#{day[0]}"
#   column = column.next
# end 
 
until column == firstemptycolumn

  row = startrow
  # If the row is a date row, set the date string to that day 
  if row = daterow
    cell = s.cell(column,daterow)
    datestring = ParseDate.parsedate cell, true
    year = datestring[0]
    month = "%02d" % datestring[1]
    day = "%02d" % datestring[2]
    date = [year,month,day].join
    row += 1
  end
  
  # Body. Grab the title string for each 30 min segment
  # If the segment is empty, or matches anything in the "info" array, consider it a continuation of the segment before
  until row > lastrow
    data = s.cell(column,row)
    start_secs = s.cell(timecolumn,row).to_i
    start_time = convert_time start_secs
    start_stamp = [ date, start_time, '00' ].join('.')
    
    if row < lastrow
      nextrow = row
      nextshow = ""
      while nextshow.empty?
        nextrow += 1
        nextshow = s.cell(column,nextrow) || ""
        unless nextshow.empty?
          if check_info info, nextshow
            data = [ data, nextshow ].join(' ')
            nextshow = ""
          end
        end
      end
      end_secs = s.cell(timecolumn,nextrow).to_i
      end_time = convert_time end_secs
      end_stamp = [ date, end_time, '00' ].join('.')
      row = nextrow
    else
      end_stamp = [ date, '240000' ].join('.')
      row += 1
    end
    
    puts "StartTime: #{start_stamp}, EndTime: #{end_stamp}, Column: #{column}, Row: #{row}, Data: '#{data}'"
      
  end
  
  column = column.next
  puts column

end
