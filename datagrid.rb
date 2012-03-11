#!/usr/bin/env ruby

require 'sinatra'
require 'spreadsheet'
require 'pry'

$:.unshift File.join(".","lib")

require 'base'

# folders
set :public_folder, File.dirname(__FILE__) + '/public'
set :files_folder, File.dirname(__FILE__) + '/files'

Spreadsheet.client_encoding = 'UTF-8'

get '/' do
  'Welcome to the reports'
end

get '/files/:year/:month/:report.:ext' do |year,month,report,ext|
  path = File.join( settings.files_folder, year, month, report ) + "." + ext
  if ext.downcase == "csv"
    send_file path
  else
    book = Spreadsheet.open path
  end
end
