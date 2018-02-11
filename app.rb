require 'bundler/setup'
Bundler.require
require 'sinatra/reloader' if development?
require "json"
require 'google/api_client/client_secrets'
require 'google/apis/sheets_v4'
require 'fileutils'
require 'open-uri'
require 'zip'
require 'rack'

enable :sessions

def user_credentials
  # Build a per-request oauth credential based on token stored in session
  # which allows us to use a shared API client.
  @authorization ||= (
  auth = settings.authorization.dup
  auth.redirect_uri = to('/oauth2callback')
  auth.update_token!(session)
  auth
  )
end

configure do
  client_secrets = Google::APIClient::ClientSecrets.load
  authorization = client_secrets.to_authorization
  authorization.scope = 'https://www.googleapis.com/auth/spreadsheets.readonly'

  set :authorization, authorization
end

before do
  # Ensure user has authorized the app
  unless user_credentials.access_token || request.path_info =~ /^\/oauth2/
    redirect to('/oauth2authorize')
  end
end

after do
  # Serialize the access/refresh token to the session and credential store.
  session[:access_token] = user_credentials.access_token
  session[:refresh_token] = user_credentials.refresh_token
  session[:expires_in] = user_credentials.expires_in
  session[:issued_at] = user_credentials.issued_at
end

get '/oauth2authorize' do
  # Request authorization
  redirect user_credentials.authorization_uri.to_s, 303
end

get '/oauth2callback' do
  # Exchange token
  user_credentials.code = params[:code] if params[:code]
  user_credentials.fetch_access_token!
  redirect to('/')
end

get '/' do
  FileUtils.rm_r("mem_pic.zip") if FileTest.exist?("mem_pic.zip")
  FileUtils.rm_r("mem_pic") if FileTest.exist?("mem_pic")
  erb :index
end

get '/save' do

  service = Google::Apis::SheetsV4::SheetsService.new
  service.client_options.application_name = "mem_pic"
  service.authorization = @authorization

  # The ID of the spreadsheet to retrieve data from.
  spreadsheet_id = params[:sheetid]  # TODO: Update placeholder value.

  # The A1 notation of the values to retrieve.
  teamrange = 'Member!B2:B356'
  namerange = 'Member!BG2:BG356'
  imgrange = 'Member!BE2:BE356'  # TODO: Update placeholder value.

  teamarray = service.get_spreadsheet_values(spreadsheet_id, teamrange).values
  namearray = service.get_spreadsheet_values(spreadsheet_id, namerange).values
  urlarray = service.get_spreadsheet_values(spreadsheet_id, imgrange).values

  teamarray.zip(namearray,urlarray) do |teams,names,urls|
    unless teams.nil? || names.nil? || urls.nil?
      teams.zip(names,urls) do |team,name,url|
        fileName = "#{team}_#{name}.jpg"
        dirName = "mem_pic"
        filePath = "#{dirName}/#{fileName}"
        FileUtils.mkdir_p(dirName) unless FileTest.exist?(dirName)

        open(filePath, 'wb') do |output|
          begin
            open(url) do |image|
              output.write(image.read)
            end
          rescue
            open("#{dirName}/#{team}_#{name}失敗.txt", 'wb').write("DL失敗")
          end
        end
      end
    end
  end
  target_files_path = "mem_pic"
  file_name = "mem_pic.zip"
  output_zip_path = "#{Dir::pwd}/#{file_name}"

  target_files = []
  Dir.glob("#{target_files_path}/*.*").each do |i|
    target_files.push(File.basename(i))
  end

  Zip::File.open(output_zip_path, Zip::File::CREATE) do |zipfile|
    target_files.each do |file|
      zipfile.add(file, "#{target_files_path}/#{file}")
    end
  end
  send_file ("mem_pic.zip")
end