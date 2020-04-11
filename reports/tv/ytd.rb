# frozen_string_literal: true

module Reports
  module Tv
    class Ytd
      def initialize(logger, driver, wait = nil)
        @logger = logger
        @driver = driver
        @wait = wait || Selenium::WebDriver::Wait.new(timeout: 300)
      end

      def run!(params = {})
        navigate_to_saved_search_tv_ytd
        pick_manual_date_range
        from_date(params[:start_date])
        to_date(params[:end_date])
        sleep(1)
        search # and download
      end

      private

      def navigate_to_saved_search_tv_ytd
        @wait.until { @driver.find_element(id: 'liSavedSearch') }
        @driver.execute_script("ChangeTabset('liSavedSearch', '/Search/')")
        @wait.until { @driver.find_element(id: 'spnQuery_276647') }
        @driver.execute_script("document.getElementById('spnQuery_276647').click();")
        @wait.until { @driver.find_element(id: 'ulSavedSearch') }
        @driver.execute_script("EditSavedSearch('liNewSearch','/Search/', 276647, 'tv YTD', false)")
      end

      def pick_manual_date_range
        # TODO: could not wait in here. must sleep
        # @wait.until { @driver.find_element(class: 'select-opener') }
        sleep(2)
        @driver.find_elements(class_name: 'select-opener').first.click
        sleep(2)
        @driver.find_elements(tag_name: 'span').select { |e| e.text == 'מותאם אישית' }.first.click
      end

      def from_date(d)
        @wait.until { @driver.find_element(id: 'FromDate') }
        @driver.execute_script("ShowDatePicker('FromDate')")
        @driver.execute_script("document.getElementById('FromDate').value='#{d}'")
      end

      def to_date(d)
        @wait.until { @driver.find_element(id: 'ToDate') }
        @driver.execute_script("ShowDatePicker('ToDate')")
        @driver.execute_script("document.getElementById('ToDate').value='#{d}'")

        # a hack to unfocus from toDate
        element = @driver.find_element(id: 'FromDate')
        element.click
      end

      def search
        # @driver.execute_script("document.getElementById('aSearch').click();")
        element = @driver.find_element(id: 'aSearch')
        element.click
      end
    end
  end
end
