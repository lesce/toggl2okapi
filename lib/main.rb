require 'rubygems'
require 'bundler/setup'

require 'csv'
require 'rubyXL'
require 'pry'
require 'ostruct'

@data = []
@duration = {}
index = 0

def extract_ticket_number task_description
  task_description.split(' ').first if task_description['FFTEC']
end

def process_duration ticket_number, date
  total_seconds = 0
  @duration["#{ticket_number}#{date}"].each do |record|
    total_seconds += 3600 * record.split(':')[0].to_i +
      60 * record.split(':')[1].to_i + 1 * record.split(':')[2].to_i
  end
  (total_seconds/3600.0).round(2)
end

# retrieve data from first csv file
CSV.foreach Dir['*.csv'].first do |row|
  index += 1
  next if index == 1 # ignore headers

  ticket_number = extract_ticket_number(row[5])
  duration      = row[11]
  date          = row[7]
  key           = "#{ticket_number}#{date}"

  if @duration[key].nil?
    @data << OpenStruct.new(
      email: row[1],
      user_name: row[0],
      task_description: row[5],
      task_number: ticket_number,
      date: date,
      duration: duration
    )
    @duration[key] = [duration]
  else
    @duration[key] << duration
  end
end

@workbook = RubyXL::Parser.parse("Template_timesheet.xlsx")
@worksheet = @workbook[0]

index = 4

@data.each do |record|
  @worksheet.insert_row(7)
end
@workbook.write("timesheet.xlsx")

@data.each do |record|
  @worksheet[index][0].change_contents(record.date)
  @worksheet[index][2].change_contents(record.user_name)
  @worksheet[index][3].change_contents(record.task_description)
  @worksheet[index][4].change_contents(process_duration(record.task_number,record.date))
  @worksheet[index][5].change_contents(record.task_number)
  index += 1
end
@workbook.write("timesheet.xlsx")
