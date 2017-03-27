"""


"""
from datetime import time
import time
import unittest
import datetime
import requests
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException
import utilities


now = datetime.datetime.now()


class TraditionalCreation(unittest.TestCase):
    """

    """
    def setUp(cls):
        """

        :return:
        """
        print 'in setup class'

        # cls.driver = webdriver.Firefox()
        # cls.driver = webdriver.Chrome()
        cls.driver = webdriver.Ie("C:\\Python27\\IEDriverServer.exe")

        cls.driver.maximize_window()
        cls.driver.implicitly_wait(40)
        wb = xlrd.open_workbook('credentials.xlsx')
        credential_sheet = wb.sheet_by_index(0)
        cec_username = credential_sheet.cell_value(0, 1)
        cec_password = credential_sheet.cell_value(1, 1)
        cls.driver.get("https://wwwin-si-stage.cisco.com/csdl/")

        ssl_certificate = cls.driver.find_element_by_xpath("//a[contains(text(),'Continue to this website (not recommended).')]")
        ssl_certificate.click()

        username_element = cls.driver.find_element_by_id("userInput")
        password_element = cls.driver.find_element_by_id("passwordInput")
        username_element.clear()
        password_element.clear()
        username_element.send_keys(cec_username)
        password_element.send_keys(cec_password)
        login_button = cls.driver.find_element_by_id("login-button")
        login_button.click()

    def call_api_siv2(self, api_call_url, header="", param=""):
        """

        :param api_call_url:
        :param header:
        :param param:
        :return:
        """
        try:
            print 'api is ',api_call_url
            requests.get(format(api_call_url), headers=header, params=param, verify=False)
        except Exception as e:
            print "Exception in ping service", str(e)

    def mandate_check(self, i, hardware_software_project, project_status, acec_date, fcs_date, ac_ec_deck_val, fcs_deck_val, psb_cserv_val, sheet1):
        """

        :param i:
        :param hardware_software_project:
        :param project_status:
        :param acec_date:
        :param fcs_date:
        :param ac_ec_deck_val:
        :param fcs_deck_val:
        :param psb_cserv_val:
        :param sheet1:
        :return:
        """
        time.sleep(3)
        self.driver.implicitly_wait("4")
        basic_mandate = []
        if hardware_software_project == '':
            utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='developing_hw_sw_project']"), 'Hardware Software Project should be mandatory', basic_mandate, sheet1)

            utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='ac_ec_date']"), 'AC/EC Date should be mandatory', basic_mandate, sheet1)

            utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='fcs_date']"), 'FCS Date should be mandatory', basic_mandate, sheet1)

        if hardware_software_project == 'Hardware':

            if project_status == 'Active' or project_status == 'Archieved':

                utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='ac_ec_date']"), 'AC/EC Date should be mandatory',basic_mandate,sheet1)
                utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='fcs_date']"), 'FCS Date should be mandatory',basic_mandate,sheet1)

        if hardware_software_project == 'Software' or hardware_software_project == 'Hardware and Software':
            utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='sw_version']"), 'Software version should be mandatory',basic_mandate,sheet1)

            if project_status == 'Active' or project_status == 'Archieved':

                utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='ac_ec_date']"), 'AC/EC Date should be mandatory',basic_mandate,sheet1)

                utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='fcs_date']"), 'FCS Date should be mandatory',basic_mandate,sheet1)

        utilities.test_for_error(self, i, 2, "//div[@data-page='traditional']/div/div[6]/ul/li[contains(text(),'Your combination of Project Name, TG/BU/PF and Software Version (if you choose Software or Hardware and Software) is not unique')]", "Your combination of Project Name, TG/BU/PF and Software Version is not unique", sheet1)

        if acec_date != "" and fcs_date != "":
            utilities.fcs_acec_check(self,i,acec_date,fcs_date,sheet1)
            if ac_ec_deck_val != "":
                utilities.test_for_error(self, i, 2, "//div[@data-field='ac_ec_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]", "Must provide URL or a valid DocCentral number", sheet1)
            if fcs_deck_val != "":
                utilities.test_for_error(self, i, 2, "//div[@data-field='fcs_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]", "Must provide URL or a valid DocCentral number", sheet1)

        elif acec_date != "":
            if ac_ec_deck_val == "":
                utilities.acec_deck_mandate(self, i, acec_date, sheet1)
            if ac_ec_deck_val != "":
                utilities.test_for_error(self, i, 2, "//div[@data-field='ac_ec_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]", "Must provide URL or a valid DocCentral number", sheet1)

        elif fcs_date != "":
            if fcs_deck_val == "":
                utilities.fcs_deck_mandate(self, i, fcs_date, sheet1)
            if fcs_deck_val != "":
                utilities.test_for_error(self,i,2,"//div[@data-field='fcs_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]","Must provide URL or a valid DocCentral number",sheet1)

        if acec_date != "" or fcs_date != "":
            todays_date = now.strftime("%d/%m/%Y")
            todays_date_date = datetime.datetime.strptime(todays_date, "%d/%m/%Y")
            if fcs_date == "":
                fcs_date = todays_date_date +datetime.timedelta(days=5)
            if acec_date == "":
                acec_date = todays_date_date +datetime.timedelta(days=5)
            if acec_date <= todays_date_date and fcs_date <= todays_date_date:
                if psb_cserv_val == "":
                    utilities.check_valid(self, i, "true", utilities.data_valid(self, "//div[@data-field='psb_evaluation']"), "PSB CSERV is mandatory", basic_mandate, sheet1)
            elif acec_date <= todays_date_date:
                if psb_cserv_val == "":
                    utilities.check_valid(self, i, "true", utilities.data_valid(self, "//div[@data-field='psb_evaluation']"), "PSB CSERV is mandatory", basic_mandate, sheet1)
            elif fcs_date <= todays_date_date:
                if psb_cserv_val == "":
                    utilities.check_valid(self, i, "true", utilities.data_valid(self, "//div[@data-field='psb_evaluation']"), "PSB CSERV is mandatory", basic_mandate, sheet1)

    def test_traditional(self):
        """

        :return:
        """
        api_url = "https://wwwin-si-stage.cisco.com/api/v1/ping/"
        sheet_index = 1
        data_set = xlrd.open_workbook('DataSet.xlsx')
        project_sheet = data_set.sheet_by_index(1)
        total_rows = project_sheet.nrows
        i = 1
        while i < total_rows:
            if i != 1:
                self.call_api_siv2(api_url)
                add_button = self.driver.find_element_by_xpath("//div[@class='add-project']")
                add_button.click()
            self.call_api_siv2(api_url)

            project_name_sheet_val, project_create_sheet = utilities.first_page(self, sheet_index, i)
            try:
                self.driver.implicitly_wait(6)
                print "Inside the ITERATION", i
                utilities.check_element_exist_sol(self, "//div[@id='csdl-nav' and @data-all-valid-on-active-page = 'true']")
                page_next_button = self.driver.find_element_by_xpath("//div[text()='Next']")
                page_next_button_validate_element = self.driver.find_element_by_xpath("//div[@id='csdl-nav']")
                page_next_button_validate_attribute = page_next_button_validate_element.get_attribute("data-all-valid-on-active-page")
                if page_next_button_validate_attribute == 'false':
                    utilities.report_error(self, project_create_sheet, i, "Some of the mandatory fields are not filled So, could not submit")
                else:
                    while page_next_button.is_displayed():
                        page_next_button.click()
                hw_sw_project_sheet_val = project_create_sheet.cell_value(i, 11)
                if hw_sw_project_sheet_val != "":
                    self.driver.implicitly_wait(15)
                    hw_sw_project_element = self.driver.find_element_by_xpath("//div[@data-field='developing_hw_sw_project']/div[3]/div[2]/fieldset/input[@value="'"' + hw_sw_project_sheet_val + '"'"]")
                    hw_sw_project_element.click()
                else:
                    hw_sw_project_sheet_val = ""

                if hw_sw_project_sheet_val == "Software":
                    if project_create_sheet.cell_value(i, 12) != "":
                        software_version_sheet_val = str(project_create_sheet.cell_value(i, 12))
                        software_version_element = self.driver.find_element_by_xpath("//div[@data-field='sw_version']/div[3]/div[2]/input")
                        software_version_element.send_keys(software_version_sheet_val)
                    else:
                        software_version_sheet_val = ""

                project_status_sheet_val = project_create_sheet.cell_value(i, 13)
                project_status_element = self.driver.find_element_by_xpath("//div[@data-field='project_status']/div[3]/div[2]/div/div/div")
                project_status_element.click()

                project_status_input_element = self.driver.find_element_by_xpath("//div[@data-field='project_status']/div[3]/div[2]/div/div[2]/input")

                project_status_input_element.send_keys(project_status_sheet_val)
                project_status_input_element.send_keys(Keys.ENTER)

                ac_ec_date_sheet_val = project_create_sheet.cell_value(i, 14)
                if ac_ec_date_sheet_val != "":
                    year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(ac_ec_date_sheet_val, data_set.datemode)

                    newacec = "%02d/%02d/%04d" % (day, month, year)
                    acec_date = datetime.datetime.strptime(newacec, "%d/%m/%Y")

                    ac_ec_date = self.driver.find_element_by_xpath("//div[@data-field='ac_ec_date']/div[3]/div[2]/div/input")
                    utilities.datepicker(self, ac_ec_date, year, month, day)
                else:
                    acec_date = ""
                fcs_date_sheet_val = project_create_sheet.cell_value(i, 15)
                if fcs_date_sheet_val != "":

                    year1, month1, day1, hour1, minute1, sec1 = xlrd.xldate_as_tuple(fcs_date_sheet_val, data_set.datemode)
                    newfcs = "%02d/%02d/%04d" % (day1, month1, year1)

                    fcs_date = datetime.datetime.strptime(newfcs, "%d/%m/%Y")

                    fcsdate = self.driver.find_element_by_xpath("//div[@data-field='fcs_date']/div[3]/div[2]/div/input")
                    utilities.datepicker(self, fcsdate, year1, month1, day1)
                else:
                    fcs_date = ""


                if project_create_sheet.cell_value(i, 16) != "":
                    ac_ec_deck_sheet_val = project_create_sheet.cell_value(i, 16)
                    ac_ec_deck = self.driver.find_element_by_xpath("//div[@data-field='ac_ec_deck']/div[3]/div[2]/input")
                    ac_ec_deck.send_keys(ac_ec_deck_sheet_val)

                else:
                    ac_ec_deck_sheet_val = ""
                if project_create_sheet.cell_value(i, 17) != "":
                    fcs_deck_sheet_val = project_create_sheet.cell_value(i, 17)
                    fcs_deck = self.driver.find_element_by_xpath("//div[@data-field='fcs_deck']/div[3]/div[2]/input")
                    fcs_deck.send_keys(fcs_deck_sheet_val)

                else:
                    fcs_deck_sheet_val = ""
                if project_create_sheet.cell_value(i, 20) != "":
                    psb_cserv_sheet_val = project_create_sheet.cell_value(i, 20)

                    psb_cserv_element = self.driver.find_element_by_xpath("//div[@data-field='psb_evaluation']/div[3]/div[2]/div/div/div/div[contains(text(),'Select a CSERV project')]")

                    psb_cserv_element.click()

                    psb_cserv_input_element = self.driver.find_element_by_xpath("//div[@data-field='psb_evaluation']/div[3]/div[2]/div/div[2]/input")

                    psb_cserv_input_element.send_keys(psb_cserv_sheet_val)
                    psb_cserv_input_element.send_keys(Keys.ENTER)
                else:
                    psb_cserv_sheet_val = ""

                self.mandate_check(i, hw_sw_project_sheet_val, project_status_sheet_val, acec_date, fcs_date, ac_ec_deck_sheet_val, fcs_deck_sheet_val,psb_cserv_sheet_val,project_create_sheet)
                self.driver.implicitly_wait(6)
                utilities.check_element_exist_sol(self, "//div[@id='csdl-nav' and @data-all-valid-on-active-page='true']")
                page_next_button = self.driver.find_element_by_xpath("//div[text()='Done']")
                validate = self.driver.find_element_by_xpath("//div[@id='csdl-nav']")

                validate_val = validate.get_attribute("data-all-valid-on-active-page")

                if validate_val == 'false':
                    utilities.report_error(self, project_create_sheet, i, "Some of the mandatory fields are not filled So, could not submit")
                else:

                    while page_next_button.is_displayed():
                        page_next_button.click()
                        print 'Done'

                    submission = "Submission Complete"
                    self.driver.implicitly_wait(30)
                    try:
                        submission_complete = self.driver.find_element_by_xpath("//div[@data-page='traditional_submitted']/h2[contains(text(),'Submission complete')]")
                    except NoSuchElementException:
                        print "Submission Incomplete"
                        submission = 'Submission Incomplete'
                        utilities.report_error(self, project_create_sheet, i, "Some thing went wrong Submission InComplete")
                    if submission == "Submission Complete":
                        utilities.report_error(self,project_create_sheet, i, "Project Submission Completed")

                        time.sleep(100)
                        self.driver.implicitly_wait(15)
                        projectTab = self.driver.find_element_by_xpath("//div[@id='subhead']/nav/ul/li[@class='projects']/a")
                        projectTab.click()
                        print 'API URL is', api_url
                        self.call_api_siv2(api_url)
                        project_name = self.driver.find_element_by_xpath("//table[@id='projects-table']/thead/tr/th[@data-field='project_name']/div[3]/input")

                        project_name.send_keys(project_name_sheet_val)

                        project_name.send_keys(Keys.ENTER)

                        if utilities.check_element_exist_sol(self,"//*[@id='projects-table']/tbody/tr/td/div/div[contains(text(),"'"'+project_name_sheet_val+'"'")]"):
                            print 'Found matching'
                            utilities.report_error(self, project_create_sheet, i, "found matching record for project")

                        else:
                            print 'No matching found'
                            utilities.report_error(self, project_create_sheet, i, "No matching record found for project")
                    else:
                        print 'Submission Incomplete for project'
                utilities.compare("Traditional Project Creation", 26, 27, 28, i)

            except NoSuchElementException as e:
                utilities.report_error(self, project_create_sheet, i, "Element Not Present")
            except ElementNotVisibleException as e:
                utilities.report_error(self, project_create_sheet, i, "Element Not Visible")
            i = i + 1

    def tearDown(cls):
        print 'inside teardownclass'
        cls.driver.quit()

if __name__ == "__main__":
    print("is being imported into another module")
    app = TraditionalCreation()
    app.run()
else:
    print("is not being imported into another module")