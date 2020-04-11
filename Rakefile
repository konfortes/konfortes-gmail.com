# frozen_string_literal: true

require 'dotenv/load'
require 'mail'
require 'logger'
require 'byebug'
require_relative './reports/report_generator'
require_relative './reports/tv/ytd'

LOG_PATH = ENV['LOG_PATH'] || 'c:\Scripts\logfile.log'
LOGIN_PATH = 'http://bakaratpirsum.co.il/Account/Login'
DOWNLOAD_PATH = ENV['DOWNLOAD_PATH'] || 'C:/Users/erank/Downloads'
TARGET_DIR_PATH = ENV['TARGET_DIR_PATH'] || '//athena/excel/Qv-YIifat/Yifat/tv current year/tv-'

namespace :utils do
  desc 'clean'
  task :clean do
    FileUtils.rm_rf(Dir.glob(DOWNLOAD_PATH + '/*'))
  end

  desc 'setup email settings'
  task :setup_mail do
    mail_options = {
      address: '192.168.50.97',
      port: 25,
      domain: 'reshet-tv.com',
      authentication: 'anonymous',
      enable_starttls_auto: true,
      openssl_verify_mode: 'none'
    }

    Mail.defaults do
      delivery_method :smtp, mail_options
    end
  end
end

namespace :tv do
  desc 'ytd'
  task ytd: ['utils:setup_mail', 'utils:clean'] do
    logger = Logger.new(LOG_PATH, 'daily')
    report_generator = Reports::ReportGenerator.new(logger, Reports::Tv::Ytd, generator_params)

    begin
      ytd_params do |start_date, end_date, name|
        report_generator.generate(start_date, end_date, name)
      end
    rescue StandardError => e
      logger.error "failed to generate report. message: #{e.message}\n trace: #{e.backtrace.join("\n")}"

      Mail.deliver do
        to 'nurit@reshet.tv'
        from 'Ifat_bakara@reshet.tv'
        subject 'Failed to generate reports for TV ' + Time.now.strftime('%H:%M - %d/%m/%Y')
        body 'see details in attached log file'
        add_file LOG_PATH
      end
      logger.info 'Email sent'
    end
  end
end

def generator_params
  {
    login_path: LOGIN_PATH,
    download_path: DOWNLOAD_PATH,
    target_dir_path: TARGET_DIR_PATH
  }
end

def ytd_params
  current_date = Date.today.strftime('%d%m%Y')

  case Date.today.strftime('%m')
  when '01', '02', '03', '04', '05', '06'
    start_date = '0101' + Date.today.strftime('%Y')
    end_date = current_date
    yield(start_date, end_date, '01-06')
  when '07', '08', '09', '10', '11', '12'
    start_date = '0101' + Date.today.strftime('%Y')
    end_date = '3006' + Date.today.strftime('%Y')
    yield(start_date, end_date, '01-06')

    start_date = '0107' + Date.today.strftime('%Y')
    end_date = current_date
    yield(start_date, end_date, '07-12')
  end
end
