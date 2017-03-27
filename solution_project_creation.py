from datetime import time
import time
import datetime
import openpyxl
import xlrd
import requests
from selenium import webdriver
import unittest
import utilities
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementNotVisibleException
from selenium.common.exceptions import UnexpectedAlertPresentException
now = datetime.datetime.now()


class SolutionCreation(unittest.TestCase):
    def setUp(cls):
        print 'in setup class'
        #cls.driver = webdriver.Firefox()
        #cls.driver.maximize_window()
        #cls.driver = webdriver.Chrome()
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

    def call_api_siv2(self,api_call_url, header="", param=""):
        try:
            print 'api is ',api_call_url
            requestObject = requests.get(format(api_call_url), headers=header, params=param, verify=False)

            # response_dict = requestObject.json()
            statusCode = int(requestObject.status_code)
            # print 'response dict is', response_dict
            print 'status code is', statusCode
        except Exception as e:
            print "Exception in ping service", str(e)

    def mandate_check(self,i,hardware_software_project,project_status,acec_date,fcs_date,ac_ec_deck_val,fcs_deck_val,psb_cserv_val,sheet1):
        time.sleep(3)
        self.driver.implicitly_wait("4")

        basic_mandate = []
        if hardware_software_project == '':
            utilities.check_valid(self, i, 'true', utilities.data_valid(self, "//div[@data-field='solution_developing_hw_sw_project']"), 'Hardware Software Project should be mandatory', basic_mandate, sheet1)

            utilities.check_valid(self, i, 'true', utilities.data_valid(self, "//div[@data-field='solution_ac_ec_date']"), 'AC/EC Date should be mandatory', basic_mandate, sheet1)

            utilities.check_valid(self, i, 'true', utilities.data_valid(self, "//div[@data-field='solution_fcs_date']"), 'FCS Date should be mandatory', basic_mandate, sheet1)

        if hardware_software_project == 'Hardware':
            if project_status == 'Active' or project_status == 'Archieved':

                utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='solution_ac_ec_date']"), 'AC/EC Date should be mandatory',basic_mandate,sheet1)
                utilities.check_valid(self, i, 'true', utilities.data_valid(self,"//div[@data-field='solution_fcs_date']"), 'FCS Date should be mandatory',basic_mandate,sheet1)

        if hardware_software_project == 'Software' or hardware_software_project == 'Hardware and Software':
            utilities.check_valid(self, i, 'true', utilities.data_valid(self, "//div[@data-field='solution_sw_version']"), 'Software version should be mandatory',basic_mandate,sheet1)

            if project_status == 'Active' or project_status == 'Archieved':

                utilities.check_valid(self, i, 'true', utilities.data_valid(self, "//div[@data-field='solution_ac_ec_date']"), 'AC/EC Date should be mandatory',basic_mandate,sheet1)

                utilities.check_valid(self, i, 'true', utilities.data_valid(self, "//div[@data-field='solution_fcs_date']"), 'FCS Date should be mandatory',basic_mandate,sheet1)
        utilities.test_for_error(self, i, 2, "//div[@data-page='solution']/div/div[4]/ul/li[contains(text(),'Your combination of Project Name, TG/BU/PF and Software Version (if you choose Software or Hardware and Software) is not unique')]", "Your combination of Project Name, TG/BU/PF and Software Version is not unique", sheet1)

        if acec_date != "" and fcs_date != "":
            utilities.fcs_acec_check(self,i,acec_date,fcs_date,sheet1)
            if ac_ec_deck_val != "":
                utilities.test_for_error(self, i, 2, "//div[@data-field='solution_ac_ec_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]", "Must provide URL or a valid DocCentral number", sheet1)
            if fcs_deck_val != "":
                utilities.test_for_error(self, i, 2, "//div[@data-field='solution_fcs_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]", "Must provide URL or a valid DocCentral number", sheet1)

        elif acec_date != "":
            if ac_ec_deck_val == "":
                utilities.acec_deck_mandate(self, i, acec_date, sheet1)
            if ac_ec_deck_val != "":
                utilities.test_for_error(self, i, 2, "//div[@data-field='solution_ac_ec_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]", "Must provide URL or a valid DocCentral number", sheet1)

        elif fcs_date != "":
            if fcs_deck_val == "":
                utilities.fcs_deck_mandate(self, i, fcs_date, sheet1)
            if fcs_deck_val != "":
                utilities.test_for_error(self,i,2,"//div[@data-field='solution_fcs_deck']/div[6]/div[@class='errors']/ul/li[contains(text(),'Must provide URL or a valid DocCentral')]","Must provide URL or a valid DocCentral number",sheet1)

        if psb_cserv_val == "":
            utilities.check_valid(self, i, 'true', utilities.data_valid(self, "//div[@data-field='solution_select_projects']"), 'Cloud or Security Insight projects should be mandatory', basic_mandate, sheet1)

    def test_solution(self):
        api_url = "https://wwwin-si-stage.cisco.com/api/v1/ping/"
        sheet_index = 2
        data_set = xlrd.open_workbook('DataSet.xlsx')
        solution_create_sheet = data_set.sheet_by_index(sheet_index)
        total_rows = solution_create_sheet.nrows
        i = 4
        print 'total_rows is', total_rows
        while i <= 5:
            if i != 4:
                print 'API URL is', api_url
                self.call_api_siv2(api_url)
                add_button = self.driver.find_element_by_xpath("//div[@class='add-project']")
                add_button.click()
                #next_button = self.driver.find_element_by_xpath("//div[@class='next']")
                #next_button.click()
            print 'API URL is', api_url
            self.call_api_siv2(api_url)

            try:
                #print 'Value of i is', i
                deployment_model_val, ProjectNametextVal, sheet1, dataSet, error_list = utilities.first_page(self, 2, i)
                self.driver.implicitly_wait(6)
                print "Inside the ITERATION", i
                utilities.check_element_exist_sol(self, "//div[@id='csdl-nav' and @data-all-valid-on-active-page = 'true']")
                page_next_button = self.driver.find_element_by_xpath("//div[text()='Next']")
                validate = self.driver.find_element_by_xpath("//div[@id='csdl-nav']")

                validate_val = validate.get_attribute("data-all-valid-on-active-page")
                if validate_val == 'false':
                    utilities.report_error(self, sheet1, i, 2, "Some of the mandatory fields are not filled So, could not submit")

                else:
                    self.driver.implicitly_wait(40)
                    while page_next_button.is_displayed():
                        page_next_button.click()
                        #print 'NEXT1'
                    hardware_software_project_sheet_val = sheet1.cell_value(i, 11)
                    if hardware_software_project_sheet_val != "":
                        hardware_software_project_element = self.driver.find_element_by_xpath("//div[@data-field='solution_developing_hw_sw_project']/div[3]/div[2]/fieldset/input[@value="'"' + hardware_software_project_sheet_val + '"'"]")
                        hardware_software_project_element.click()
                    else:
                        hardware_software_project_sheet_val = ""
                    if hardware_software_project_sheet_val == "Software":
                        software_version_sheet_val = sheet1.cell_value(i, 12)
                        if software_version_sheet_val != "":
                            software_version = self.driver.find_element_by_xpath("//div[@data-field='solution_sw_version']/div[3]/div[2]/input")
                            software_version.send_keys(software_version_sheet_val)
                        else:
                            software_version_sheet_val = ""

                    project_status_element = self.driver.find_element_by_xpath("//div[@data-field='solution_project_status']/div[3]/div[2]/div/div/div")
                    project_status_element.click()

                    project_status_sheet_value = sheet1.cell_value(i, 13)
                    self.driver.find_element_by_xpath("//div[@data-field='solution_project_status']/div[3]/div[2]/div/div[2]/input")
                    if project_status_sheet_value != "":
                        project_status_input = self.driver.find_element_by_xpath("//div[@data-field='solution_project_status']/div[3]/div[2]/div/div[2]/input")
                        project_status_input.send_keys(project_status_sheet_value)
                        project_status_input.send_keys(Keys.ENTER)
                    else:
                        project_status_sheet_value = "Active"
                    ac_ec_date_sheet_val = sheet1.cell_value(i, 14)
                    self.driver.find_element_by_xpath("//div[@data-field='solution_ac_ec_date']/div[3]/div[2]/div/input")
                    if ac_ec_date_sheet_val != "":
                        year, month, day, hour, minute, sec = xlrd.xldate_as_tuple(ac_ec_date_sheet_val, dataSet.datemode)

                        newacec = "%02d/%02d/%04d" % (day, month, year)
                        acec_date = datetime.datetime.strptime(newacec, "%d/%m/%Y")

                        ac_ec_date_element = self.driver.find_element_by_xpath("//div[@data-field='solution_ac_ec_date']/div[3]/div[2]/div/input")
                        utilities.datepicker(self, ac_ec_date_element, year, month, day)
                    else:
                        acec_date = ""

                    fcs_date_sheet_val = sheet1.cell_value(i, 15)
                    self.driver.find_element_by_xpath("//div[@data-field='solution_fcs_date']/div[3]/div[2]/div/input")
                    if fcs_date_sheet_val != "":

                        year1, month1, day1, hour1, minute1, sec1 = xlrd.xldate_as_tuple(fcs_date_sheet_val, dataSet.datemode)
                        newfcs = "%02d/%02d/%04d" % (day1, month1, year1)

                        fcs_date = datetime.datetime.strptime(newfcs, "%d/%m/%Y")

                        fcsdate = self.driver.find_element_by_xpath("//div[@data-field='solution_fcs_date']/div[3]/div[2]/div/input")
                        utilities.datepicker(self, fcsdate, year1, month1, day1)
                    else:
                        fcs_date = ""

                    ac_ec_deck_sheet_value = sheet1.cell_value(i, 16)
                    self.driver.find_element_by_xpath("//div[@data-field='solution_ac_ec_deck']/div[3]/div[2]/input")
                    if ac_ec_deck_sheet_value != "":
                        ac_ec_deck = self.driver.find_element_by_xpath("//div[@data-field='solution_ac_ec_deck']/div[3]/div[2]/input")
                        ac_ec_deck.send_keys(ac_ec_deck_sheet_value)

                    else:
                        ac_ec_deck_sheet_value = ""
                    fcs_deck_sheet_val = sheet1.cell_value(i, 17)
                    self.driver.find_element_by_xpath("//div[@data-field='solution_fcs_deck']/div[3]/div[2]/input")
                    if fcs_deck_sheet_val != "":
                        fcs_deck = self.driver.find_element_by_xpath("//div[@data-field='solution_fcs_deck']/div[3]/div[2]/input")
                        fcs_deck.send_keys(fcs_deck_sheet_val)

                    else:
                        fcs_deck_sheet_val = ""
                    security_insight_project_sheet_val = sheet1.cell_value(i, 19)
                    self.driver.find_element_by_xpath("//div[@data-field='solution_select_projects']/div[3]/div[2]/div/div/input")
                    if security_insight_project_sheet_val != "":

                        security_insight_project_element = self.driver.find_element_by_xpath("//div[@data-field='solution_select_projects']/div[3]/div[2]/div/div/input")

                        security_insight_project_element.click()
                        security_insight_project_element.send_keys(security_insight_project_sheet_val)
                        security_insight_project_element.send_keys(Keys.ENTER)
                        print 'security_insight_project_sheet_val is', security_insight_project_sheet_val
                    else:
                        security_insight_project_sheet_val = ""
                    self.mandate_check(i, hardware_software_project_sheet_val, project_status_sheet_value, acec_date, fcs_date, ac_ec_deck_sheet_value, fcs_deck_sheet_val, security_insight_project_sheet_val, sheet1)

                    page_done_button = self.driver.find_element_by_xpath("//div[text()='Done']")
                    validate = self.driver.find_element_by_xpath("//div[@id='csdl-nav']")
                    time.sleep(3)
                    validate_val = validate.get_attribute("data-all-valid-on-active-page")
                    #print 'validate is ', validate_val
                    if validate_val == 'false':
                        utilities.report_error(self, sheet1, i, 2, "Some of the mandatory fields are not filled So, could not submit")
                    else:
                        self.driver.implicitly_wait(30)
                        while page_done_button.is_displayed():
                            page_done_button.click()
                            print 'Done'
                        submission = "Submission Complete"
                        try:
                            submission_complete = self.driver.find_element_by_xpath("//div[@data-page='solution_submitted']/h2[contains(text(),'Submission complete')]")
                        except NoSuchElementException:
                            print "Submission Incomplete"
                            submission = 'Submission Incomplete'
                            utilities.report_error(self, sheet1, i, 2, "Something went wrong Submission InComplete")
                        if submission == "Submission Complete":
                            utilities.report_error(self, sheet1, i, 2, "Submission Completed")
                            #self.driver.implicitly_wait(60)
                            time.sleep(60)
                            self.driver.implicitly_wait(15)
                            solution_tab = self.driver.find_element_by_xpath("//div[@id='subhead']/nav/ul/li[@class='solutions']/a")
                            solution_tab.click()
                            print 'API URL is', api_url
                            self.call_api_siv2(api_url)
                            solution_name = self.driver.find_element_by_xpath("//table[@id='solutions-table']/thead/tr/th[@data-field='solution_name']/div[3]/input")
                            # Project_name_static = "Test Project 11"
                            solution_name.send_keys(ProjectNametextVal)
                            time.sleep(2)
                            solution_name.send_keys(Keys.ENTER)
                            time.sleep(2)
                            if utilities.check_element_exist(self, "//*[@id='solutions-table']/tbody/tr/td/div/div[contains(text(),"'"'+ProjectNametextVal+'"'")]"):
                                print 'Found matching'
                                utilities.report_error(self, sheet1, i, 2, "Found matching record for solution")
                            else:
                                print 'No matching found'
                                utilities.report_error(self, sheet1, i, 2, "No Found matching record for solution")
                        else:
                            print 'Submission Incomplete for solution ' + ProjectNametextVal + ' '
                utilities.compare("Solution Project Creation", 20, 21, 22, i)
            except NoSuchElementException as e:
                print 'Element Not Present', str(e)
            except ElementNotVisibleException as e:
                print 'Element Not Visible', str(e)
                #utilities.report_error(self, sheet1, i, 2, "Element Not Present")
            i = i+1

    def tearDown(cls):
        print 'inside teardownclass'
        cls.driver.quit()

