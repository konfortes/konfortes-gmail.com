# frozen_string_literal: true

require 'selenium-webdriver'
require 'json'
require 'net/http'
require 'fileutils'
require 'date'

module Reports
  class ReportGenerator
    def initialize(logger, report_class, params, driver = create_driver)
      @logger = logger
      @params = params
      @driver = driver
      @wait = Selenium::WebDriver::Wait.new(timeout: 300) # in seconds
      @report = report_class.new(@logger, @driver, @wait)
    end

    def generate(start_date, end_date, name)
      @logger.debug "start generating report #{name} from #{start_date} to #{end_date}"

      login
      disable_popups
      @driver.switch_to.default_content
      @report.run!(start_date: start_date, end_date: end_date, name: name)

      if @driver.find_element(tag_name: 'body').text.include?('לא נמצאו תוצאות מתאימות להגדרת החיפוש') == true
        @logger.info('no search results - no file is downloaded') && return
      end

      copy_downloaded_report(name)
    ensure
      @driver.quit
    end

    private

    def create_driver
      capabilities = Selenium::WebDriver::Remote::Capabilities.firefox(marionette: true)

      profile = Selenium::WebDriver::Firefox::Profile.new
      profile['browser.download.folderList'] = 2
      profile['browser.download.saveLinkAsFilenameTimeout'] = 1
      profile['browser.download.manager.showWhenStarting'] = false
      profile['browser.download.dir'] = 'c:/temp/test'
      profile['browser.download.downloadDir'] = 'c:/temp/test'
      profile['browser.download.defaultFolder'] = 'c:/temp/test'
      profile['browser.helperApps.neverAsk.saveToDisk'] = 'text/csv'
      profile['plugin.scan.plid.all'] = false

      Selenium::WebDriver.for :firefox, desired_capabilities: capabilities, profile: profile
    end

    def login
      @driver.navigate.to(@params[:login_path])

      @wait.until { @driver.find_element(id: 'userName') }
      element = @driver.find_element(id: 'userName')
      element.clear
      element.send_keys(ENV['LOGIN_EMAIL'])

      element = @driver.find_element(name: 'password')
      element.clear
      element.send_keys(ENV['LOGIN_PASSWORD'])

      @driver.execute_script("document.getElementById('btnSubmit').click();")

      @wait.until { @driver.find_element(id: 'divHold') != nil }
      @logger.info 'logged in successfully'
    end

    def disable_popups
      @driver.manage.add_cookie(name: 'picreel_popup__passed', value: '1533047', path: '/', domain: 'bakaratpirsum.co.il')
      @driver.manage.add_cookie(name: 'picreel_popup__viewed', value: '1579477', path: '/', domain: 'bakaratpirsum.co.il')
      @driver.manage.add_cookie(name: 'picreel_popup__template_passed_1579477', value: '1579477', path: '/', domain: 'bakaratpirsum.co.il')
      @driver.manage.add_cookie(name: 'picreel_popup__template_passed_1579590', value: '1579590', path: '/', domain: 'bakaratpirsum.co.il')

      @logger.info 'planted cookies to disable popup'
    end

    def copy_downloaded_report(name)
      download_path_pattern = @params[:download_path] + '/*'
      sleep(35)
      # @wait.until { Dir.glob(download_path_pattern).none? { |x| x.include? '.part' } }

      current_file = Dir.glob(download_path_pattern).max_by { |f| File.mtime(f) }
      raise 'error downloading report' if current_file.nil?

      @logger.info "Report File successfully downloaded to #{@params[:download_path]}"
      FileUtils.cp_r(current_file, "#{@params[:target_dir_path]}#{name}.csv")
      @logger.info "report file successfully copied to #{@params[:target_dir_path]}"
    end
  end
end
