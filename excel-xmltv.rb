require 'rubygems'
#require 'spreadsheet'
require 'roo'
require 'parsedate'
require 'builder'
require './lib/program_info'

# Extra Homework: Find out fill colour for additional program data
# 'spreadsheet' gem test
# book = Spreadsheet.open 'triangle.xls'
# sheet = book.worksheet 0
# 
# sheet.each do |row|
#   puts row.join(',')
#   format = row.format 3
#   # puts format.to_a
# end

daterow = 2
startrow = 3
lastrow = 52
startcolumn = "b"
firstemptycolumn = "i"
timecolumn = "a"
file = "triangle.xls"
column = startcolumn

# Channel info for XML
channel_id = "tritv.co.nz"
display_name = "Triangle TV"
icon_src = "http://www.littledayout.co.nz/images/TRI-TV_200px.jpg"
default_rating = "G"
rating_system = "Freeview"

# Regexes for rating and year tags
rating_regex = /\[[A-Z]+\]/
year_regex = /([0-9]{4})/

# Excel file handle
s = Excel.new(file)
s.default_sheet = s.sheets.first

# XML builder init
x = Builder::XmlMarkup.new(:target=>STDOUT, :indent=>2)
shows = Hash.new
show_id = 1

# Get an xmltv compliant date string
def date_string(date)
  date_string = ParseDate.parsedate date
  year = date_string[0]
  month = "%02d" % date_string[1]
  day = "%02d" % date_string[2]
  [year,month,day].join
end

# Convert date and seconds into xmltv time string
def convert_time(date, secs)
  mins = secs / 60
  hours = ((mins % (60 *24)) / 60).floor
  hours = "%02d" % hours
  minutes = mins % 60
  minutes = "%02d" % minutes
  [ date, hours, minutes, '00'].join
end

until column == firstemptycolumn

  row = startrow
  # If the row is a date row, set the date variable to that day 
  if row = daterow
    cell = s.cell(column,daterow)
    date = date_string cell
    row += 1
  end
  
  # Body. Grab the title string for each 30 min segment
  # If the segment is empty, or matches anything in the "info" array, consider it a continuation of the segment before
    
  until row > lastrow
    data = s.cell(column,row)
    start_secs = s.cell(timecolumn,row).to_i
    start_time = convert_time date, start_secs
    
    if row < lastrow
      nextrow = row
      nextshow = ""
      rating = ""
      year = ""
      while nextshow.empty?
        nextrow += 1
        nextshow = s.cell(column,nextrow) || ""
        unless nextshow.empty?
          # This could be a rating or year tag. Checking here
          pi = ProgInfo.new(rating_regex, year_regex, nextshow)
          
          rating_find = pi.rating
          year_find = pi.year
          
          rating = rating_find unless rating_find.empty?
          year = year_find unless year_find.empty?
          
          nextshow = "" if rating_find != "" || year_find != ""
          
        end
      end
      
      # If no rating was found, set it to the default rating
      rating = default_rating if rating.empty?
      # Shamelessly hack a month and day onto the end of the year string
      year = [ year, "0101"].join unless year.empty?
      end_secs = s.cell(timecolumn,nextrow).to_i
      end_time = convert_time date, end_secs
      row = nextrow
      
    else
      end_stamp = [ date, '240000' ].join
      row += 1
    end
    
    # Push the found show info into a hash in the shows array
    show = Hash['starttime' => start_time ,
                'endtime' => end_time,
                'data' => data,
                'rating' => rating,
               ]

    # If a year was found, push it onto the end of the show hash
    show['year'] = year unless year.empty?

    # This is an array ID hack. Just to get nice sorting on the shows array
    shows[show_id] = show
    show_id += 1
    #puts show.inspect
      
  end

  # Move onto the next column
  column = column.next

end

# Generate a timestamp for now to show when the XML was generated
t = Time.now
now_time = [ t.year, t.mon, t.day, t.hour, t.min, t.sec ].join.to_s

# Build the XML
x.tv("date" => now_time, "generator-info-name" => "excel-xmltv") {
  
  # XML channel crap for standard top tags
  x.channel("id" => channel_id) {
    x.tag! 'display-name', display_name
    x.icon("src" => icon_src)
  }
  
  # XML show data
  shows.sort.each do |show|
    #puts show.inspect
    starttime = show[1]['starttime']
    endtime = show[1]['endtime']
    data = show[1]['data']
    rating = show[1]['rating']
    year = show[1]['year']
    
    x.programme("channel" => channel_id, "start" => starttime, "stop" => endtime) {
      x.title(data)
      if year
        x.date(year)
      end
      x.rating("system" => rating_system) {
        x.value(rating)
      }
    }
  end
}
