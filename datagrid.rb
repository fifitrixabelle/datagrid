#!/usr/bin/env ruby

require 'sinatra'
require 'csv'
require 'spreadsheet'
require 'pry'

# folders
set :public_folder, File.dirname(__FILE__) + '/public'
set :files_folder, File.dirname(__FILE__) + '/files'

Spreadsheet.client_encoding = 'UTF-8'

get '/' do
  'Welcome to the reports'
end

get '/files/:year/:month/:report.:ext' do |year,month,report,ext|
  path = File.join( settings.files_folder, year, month, report ) + "." + ext
  @report = report
  case ext.downcase
  when 'html'
    @csv = CSV.read( path.sub( /html$/, 'csv' ), :encoding => 'windows-1251:utf-8' )
    return erb :grid
  when 'xls'
    if ! File.exists? path
      render_xls path
    end
  end
  send_file path
end

private

def render_xls( path )
  csv = CSV.read( path.sub( /xls$/, 'csv' ), :encoding => 'windows-1251:utf-8' )

  save_path = path
  book = Spreadsheet::Workbook.new
  sheet = book.create_worksheet
  sheet.name = @report

  #write header
  sheet.row(0).concat csv.shift

  #write data
  y=1
  csv.each do |row|
    x=0
    row.each do |cell|
      sheet[y,x] = cell
      x+=1
    end
    y+=1
  end

  #formatting
  sheet.row(0).height = 18
  format = Spreadsheet::Format.new :color => :blue,
                                   :weight => :bold,
                                   :size => 18
  sheet.row(0).default_format = format

  book.write save_path
end
