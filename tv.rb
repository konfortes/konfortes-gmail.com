require 'selenium-webdriver'
require 'logger'
require 'json'
require 'net/http'
require 'fileutils'
require 'date'
require 'mail'
require 'ntlm/smtp'

LOGIN_PATH = 'http://bakaratpirsum.co.il/Account/Login'.freeze
DOWNLOAD_PATH = 'C:/Users/erank/Downloads'.freeze
TARGET_DIR = '//athena/excel/Qv-YIifat/Yifat/tv current year/tv-'.freeze
LOG_PATH = 'c:\Scripts\logfile.log'.freeze

$logger = Logger.new(LOG_PATH, 'daily')

$logger.info '---------------------------------------------------- TV Start---------------------------------------------------------'

class Tv
  def initialize(driver)
    @driver = driver
    @wait = Selenium::WebDriver::Wait.new(timeout: 480) # in seconds
  end

  def generate(start_date, end_date, name)
    $logger.debug "start date is #{start_date}"

    login
    disable_popups
    @driver.switch_to.default_content

    saved_search_tv_ytd

    manual_date_range
    from_date(start_date)
    to_date(end_date)
    sleep(2)
    search # and download

    copy_downloaded_report

    sleep(2)
    @driver.quit
  end

  private

  def login
    @driver.navigate.to(LOGIN_PATH)

    @wait.until { @driver.find_element(id: 'userName') }
    element = @driver.find_element(id: 'userName')
    element.clear
    element.send_keys("iritk@reshet.tv")

    element = @driver.find_element(name: 'password')
    element.clear
    element.send_keys('iritk')

    # TODO: remove?
    # element = @driver.find_element(id: 'btnSubmit')
    @driver.execute_script("document.getElementById('btnSubmit').click();")

    @wait.until { @driver.find_element(id: 'divHold') != nil }
    $logger.info "logged in successfully to " + name + " report"
  end

  def disable_popups
    @driver.manage.add_cookie(name: 'picreel_popup__passed', value: '1533047', path: '/', domain: 'bakaratpirsum.co.il');
    @driver.manage.add_cookie(name: 'picreel_popup__viewed', value: '1579477', path: '/', domain: 'bakaratpirsum.co.il');
    @driver.manage.add_cookie(name: 'picreel_popup__template_passed_1579477' , value: '1579477', path: '/', domain: 'bakaratpirsum.co.il');
    @driver.manage.add_cookie(name: 'picreel_popup__template_passed_1579590' , value: '1579590', path: '/', domain: 'bakaratpirsum.co.il');

    $logger.info 'planted cookies to disable popup'
  end

  def saved_search_tv_ytd
    sleep(2)
    @driver.execute_script("ChangeTabset('liSavedSearch', '/Search/')")
    sleep(2)
    @driver.execute_script("document.getElementById('spnQuery_276647').click();") 
    sleep(2)
    @driver.execute_script("EditSavedSearch('liNewSearch','/Search/', 276647, 'tv YTD', false)")
    sleep(2)
  end

  def manual_date_range
    @driver.find_elements(:class_name => "select-opener").first.click 
    sleep(1)
    @driver.find_elements(:tag_name => "span").select {|e| e.text == "מותאם אישית"}.first.click
    sleep(2)
  end

  def from_date(d)
    @driver.execute_script("ShowDatePicker('FromDate')") 
    element = @driver.find_element(id: 'FromDate')
    element.send_keys(d)
  end

  def to_date(d)
    @driver.execute_script("ShowDatePicker('ToDate')") 
    element = @driver.find_element(id: 'ToDate')
    element.send_keys(d)
  end

  def search
    @driver.execute_script("document.getElementById('aSearch').click();")
    sleep(25)
  end

  def copy_downloaded_report
    if @driver.find_element(tag_name: 'body').text.include?('לא נמצאו תוצאות מתאימות להגדרת החיפוש') == true
      $logger.info 'no search results - no file is downloaded'
    else
      $logger.info 'search results - success - file is being downloaded'
      wait = Selenium::WebDriver::Wait.new(timeout: 540) # in seconds
      sleep(2)
      wait.until { Dir.glob(DOWNLOAD_PATH + '/*').any? { |x| x.include? '.part' } == false }
      $logger.info 'Report File successfully downloaded'

      # TODO: UNCOMMENT
      # currect_file = Dir.glob(DOWNLOAD_PATH + '/*').max_by { |f| File.mtime(f) }
      # FileUtils.cp_r(currect_file, TARGET_DIR + name + '.csv')
      $logger.info 'report file created successfully for ' + name
    end
  end
end

def create_driver
  capabilities = Selenium::WebDriver::Remote::Capabilities.firefox(marionette: false)

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

def setup_mail
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

def execute
  report_generator = Tv.new(create_driver)
  current_date = Date.today.strftime('%d%m%Y')

  begin
    case Date.today.strftime('%m')
    when '01', '02', '03', '04', '05', '06'
      $logger.info 'started creating reports for first half year - up to June inclusive'
      start_date = '0101' + Date.today.strftime('%Y')
      end_date = current_date
      report_generator.generate(start_date, end_date, '01-06')
    when '07', '08', '09', '10', '11', '12'
      $logger.info 'started creating reports for all year devided by up to June and July forth'
      report_generator.generate('0101' + Date.today.strftime('%Y'), '3006' + Date.today.strftime('%Y'), '01-06')
      start_date = '0107' + Date.today.strftime('%Y')
      end_date = current_date
      report_generator.generate(start_date, end_date, '07-12')
    end
  rescue StandardError => e
    $logger.error "failed to run (in rescue) #{e.message}"
    setup_mail
    Mail.deliver do
      to 'nurit@reshet.tv'
      from 'Ifat_bakara@reshet.tv'
      subject 'Failed to generate reports for TV ' + Time.now.strftime('%H:%M - %d/%m/%Y')
      body 'see details in attached log file'
      add_file 'c:\Scripts\logfile.log'
    end
    $logger.info 'Email sent'
  end
end

FileUtils.rm_rf(Dir.glob(DOWNLOAD_PATH + '/*'))
execute
