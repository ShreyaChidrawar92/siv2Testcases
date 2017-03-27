"""
This module helps to checks different scenarios while creating cloud_projects
"""
from datetime import time
import time
import unittest
import requests
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException
import utilities


class CloudCreation(unittest.TestCase):

    """
     class to create cloud project
    """

    def setUp(self):

        """
        This method is used t
        """

        # self.driver = webdriver.Firefox()
        self.driver = webdriver.Chrome()
        # self.driver = webdriver.Ie("C:\\Python27\\IEDriverServer.exe")
        self.driver.get("https://wwwin-si-stage.cisco.com/csdl/")
        # ssl_certificate = self.driver.find_element_by_xpath("//a[contains(text(),'Continue to this website (not recommended).')]")
        # ssl_certificate.click()
        self.driver.maximize_window()
        self.driver.implicitly_wait(40)
        wb = xlrd.open_workbook('credentials.xlsx')
        credential_sheet = wb.sheet_by_index(0)
        cec_username = credential_sheet.cell_value(0, 1)
        cec_password = credential_sheet.cell_value(1, 1)
        self.driver.get("https://wwwin-si-stage.cisco.com/csdl/")
        username_element = self.driver.find_element_by_id("userInput")
        password_element = self.driver.find_element_by_id("passwordInput")
        username_element.clear()
        password_element.clear()
        username_element.send_keys(cec_username)
        password_element.send_keys(cec_password)
        login_button = self.driver.find_element_by_id("login-button")
        login_button.click()

    def call_api_siv2(self,api_call_url, header="", param=""):
        try:
            requests.get(format(api_call_url), headers=header, params=param, verify=False)
        except Exception as e:
            print "Exception in ping service", str(e)

    def test_cloud(self):
        """
        this method is to iterate on all the rows(scenario's) provided in excel
        :return:
        """
        api_url = "https://wwwin-si-stage.cisco.com/api/v1/ping/"
        sheet_index = 0
        data_set = xlrd.open_workbook('DataSet.xlsx')
        cloud_create_sheet = data_set.sheet_by_index(sheet_index)
        total_rows = cloud_create_sheet.nrows
        i = 1
        while i < total_rows:
            if i != 1:
                self.call_api_siv2(api_url)
                add_button = self.driver.find_element_by_xpath("//div[@class='add-project']")
                add_button.click()
            self.call_api_siv2(api_url)

            project_name_sheet_val, cloud_create_sheet = utilities.first_page(self, sheet_index, i)
            try:
                self.driver.implicitly_wait(10)
                print "Inside the ITERATION", i
                utilities.check_element_exist_sol(self,"//div[@id='csdl-nav' and @data-all-valid-on-active-page = 'true']")
                page_next_button = self.driver.find_element_by_xpath("//div[text()='Next']")
                page_next_button_validate_element = self.driver.find_element_by_xpath("//div[@id='csdl-nav']")
                page_next_button_validate_attribute = page_next_button_validate_element.get_attribute("data-all-valid-on-active-page")

                if page_next_button_validate_attribute == 'false':
                    utilities.report_error(self, cloud_create_sheet, i, "Some of the mandatory fields are not filled So, could not submit")
                else:
                    while page_next_button.is_displayed():
                        page_next_button.click()
                        time.sleep(1)
                    submission = 'Submission Complete'
                    try:
                        self.driver.implicitly_wait(60)
                        self.driver.find_element_by_xpath("//div[@data-page='cloud_submitted']/h2[contains(text(),'Submission complete to CSDL')]")
                    except NoSuchElementException:
                        submission = 'Submission Incomplete'
                        utilities.report_error(self, cloud_create_sheet, i, "Project Submission Failed")

                    if submission == "Submission Complete":
                        utilities.report_error(self, cloud_create_sheet, i, "Project Submission Completed")
                        time.sleep(50)
                        self.driver.implicitly_wait(15)
                        cloud_tab = self.driver.find_element_by_xpath("//*[@id='subhead']/nav/ul/li/a[text()='Cloud']")
                        cloud_tab.click()
                        self.call_api_siv2(api_url)
                        cloud_project = self.driver.find_element_by_xpath("//*[@id='cloud-table']/thead/tr/th[1]/div[3]/input")
                        cloud_project.send_keys(project_name_sheet_val)
                        cloud_project.send_keys(Keys.ENTER)
                        if utilities.check_element_exist_sol(self, "//*[@id='cloud-table']/tbody/tr/td/div/div[@class='project_name' and contains(text(),"'"' + project_name_sheet_val + '"'")]"):
                            utilities.report_error(self, cloud_create_sheet, i, "Matching record found for project")

                        if utilities.check_element_exist_sol(self, "//*[@id='cloud-table']/tbody/tr/td[contains(text(),'No matching')]"):
                            utilities.report_error(self, cloud_create_sheet, i, "No Matching record found")

                utilities.compare("Cloud Project Creation", 11, 12, 13, i)
                i = i + 1
            except NoSuchElementException as e:
                utilities.report_error(self, cloud_create_sheet, i, "No Such Element Present "+str(e)+" ")
                print 'Element Not Present', str(e)
            except ElementNotVisibleException as e:
                utilities.report_error(self, cloud_create_sheet, i, "Element Not Visible "+str(e)+" ")
                print 'Element Not Visible', str(e)

    def tearDown(self):
        self.driver.quit()














