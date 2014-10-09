"""
    Remote ranorex library for robot framework
    All commands return True if they are executed correctly
"""
#iron python ]%imports
import clr
clr.AddReference('Ranorex.Core')
clr.AddReference('System.Windows.Forms')
import System.Windows.Forms
import Ranorex
#python imports
from argparse import ArgumentParser
from robotremoteserver import RobotRemoteServer
import subprocess
import logging
import time
import sys
import os

log = logging.getLogger("RXCONNECTOR")
timeOut = 10000 # 10 seconds
class RanorexLibrary(object):
    """ Basic implementation of ranorex object calls for
    robot framework
    """
    def __init__(self):
        self.debug = False
        self.model_loaded = False
        self.model = None
        Ranorex.Mouse.DefaultMoveTime = 0 # default 300 miliseconds
        Ranorex.Keyboard.DefaultKeyPressTime = 5 # default 100 miliseconds
        Ranorex.Delay.SpeedFactor = 0
        Ranorex.Adapter.DefaultSearchTimeout = 1000 # 1 second


    @classmethod
    def __return_type(cls, locator):
        """ Function serves as translator from xpath into
        .net object that is recognized by ranorex.
        Returns supported object type.
        """
        Ranorex.Validate.EnableReport = False
        Ranorex.Adapter.DefaultUseEnsureVisible = True
        supported_types = ['AbbrTag', 'AcronymTag', 'AddressTag', 'AreaTag',
                           'ArticleTag', 'AsideTag', 'ATag', 'AudioTag',
                           'BaseFontTag', 'BaseTag', 'BdoTag', 'BigTag',
                           'BodyTag', 'BrTag', 'BTag', 'Button',
                           'ButtonTag', 'CanvasTag', 'Cell', 'CenterTag',
                           'CheckBox', 'CiteTag', 'CodeTag', 'ColGroupTag',
                           'ColTag', 'Column', 'ComboBox', 'CommandTag',
                           'Container', 'ContextMenu', 'DataListTag',
                           'DdTag', 'DelTag', 'DetailsTag', 'DfnTag',
                           'DirTag', 'DivTag', 'DlTag', 'EmbedTag', 'EmTag',
                           'FieldSetTag', 'FigureTag', 'FontTag', 'Form', 'FormTag',
                           'Link', 'List', 'ListItem', 'MenuBar',
                           'MenuItem', 'Picture', 'ProgressBar',
                           'RadioButton', 'Row', 'ScrollBar', 'Slider',
                           'StatusBar', 'Table', 'Text', 'TitleBar', 'ToggleButton',
                           'Tree', 'TreeItem', 'Unknown','TabPage']
        splitted_locator = locator.split('/')
        if "[" in splitted_locator[-1]:
            ele = splitted_locator[-1].split('[')[0]
        else:
            ele = splitted_locator[-1]
        if ele.lower() == '':
                raise AssertionError("No element entered")
        for item in supported_types:
            if ele.lower() == item.lower():
                return item
        raise AssertionError("Element type is not supported. Entered element: %s" %
                             ele)

    def __create_element(self, locator):
        element_created = False
        try:
            element_type = None
            element_type = self.__return_type(locator)
            element = getattr(Ranorex, element_type)(locator)
            element_created = True
            if self.debug:
                log.debug("Element at %s", locator)
                log.debug("Application object is %s", element)
        except:
            log.debug("Element %s not found!", element_type)
        if not element_created:
            raise AssertionError("Element %s not found!" %element_type)

        return element

    def rn_start_debug(self):
        """ Starts to show debug messages on remote connector """
        self.debug = True

    def rn_stop_debug(self):
        """ Stops to show debug messages """
        self.debug = False

    def rn_click_element(self, locator, location=None):
        """ Clicks on element identified by locator and location

        :param locator: xpath selector of element
        :param location: relative coordinates of mouse click from
                         top left corner of element, i.e. "x,y"

        :returns: True / False
        """

        if self.debug:
            log.debug("Click Element")
            log.debug("Location: %s", location)
        element = self.__create_element(locator)
        try:
            if location == None:
                element.Click()
                return True
            else:
                if not isinstance(location, basestring):
                    raise AssertionError("Location must be a string")
                location = [int(x) for x in location.split(',')]
                element.Click(Ranorex.Location(location[0], location[1]))
                return True
        except Exception as error:
            if self.debug:
                log.error("Failed because of %s", error)
            raise AssertionError(error)

    def rn_check(self, locator):
        """ Check if element is checked. If not it check it.
            Only checkbox and radiobutton are supported.
            Uses Click() method to check it.

        :param locator: xpath selector of element
        :returns: True/False
        """

        if self.debug:
            log.debug("Check")
        element = self.__create_element(locator)
        if not element.Element.GetAttributeValue('Checked'):
            element.Click()
        return True

    @classmethod
    def rn_check_if_process_is_running(cls, process_name):
        """ Check if process with desired name is running.
            Returns name of process if running

        :param process_name: xpath selector of element
        :returns: True/False
        """

        proc = subprocess.Popen(['tasklist'], stdout=subprocess.PIPE)
        out = proc.communicate()[0]
        return out.find(process_name) != -1 if out else False

    def rn_clear_text(self, locator):
        """ Clears text from text box. Only element Text is supported.

        :param locator: xpath selector of element
        :returns: True/False
        """

        if self.debug:
            log.debug("Clear Text")
        element = self.__create_element(locator)
        element.PressKeys("{End}{Shift down}{Home}{Shift up}{Delete}")
        return True

    def rn_double_click_element(self, locator, location=None):
        """ Doubleclick on element identified by locator. It can click
            on desired location if requested.

        :param locator: xpath selector of element
        :param location: relative coordinates of mouse click from
                         top left corner of element, i.e. "x,y"

        :returns: True/False
        """

        if self.debug:
            log.debug("Double Click Element")
            log.debug("Location: %s", location)
        element = self.__create_element(locator)
        try:
            if location == None:
                element.DoubleClick()
                return True
            else:
                if not isinstance(location, basestring):
                    raise AssertionError("Location must be a string")
                location = [int(x) for x in location.split(',')]
                element.DoubleClick(Ranorex.Location(location[0], location[1]))
                return True
        except Exception as error:
            raise AssertionError(error)

    def rn_get_table(self, locator):
        """ Get content of table without headers

        :param locator: xpath string selecting element on screen
        :returns: two dimensional array with content of the table
        """

        element = self.__create_element(locator)
        table = [[cell.Text for cell in row.Cells] for row in element.Rows]
        return table

    def rn_get_element_attribute(self, locator, attribute):
        """ Get specified element attribute.

        :param locator: xpath selector of element
        :returns: True/False
        """
        if self.debug:
            log.debug("Get Element Attribute %s", attribute)
        element = self.__create_element(locator)
        _attribute = element.Element.GetAttributeValue(attribute)
        if self.debug:
            log.debug("Found attribute value is: %s", _attribute)
        return _attribute

    def rn_input_text(self, locator, text):
        """ input texts into specified locator.

        :param locator: xpath selector of element
        :param text: text value to input into element
        :returns: True/False
        """
        self.rn_clear_text(locator)
        if self.debug:
            log.debug("Input Text: %s", text)
        element = self.__create_element(locator)
        element.PressKeys(text)
        return True

    def rn_right_click_element(self, locator, location=None):
        """ Rightclick on desired element identified by locator.
        Location of click can be used.

        :param locator: xpath selector of element
        :param location: relative coordinates of mouse click from
                         top left corner of element, i.e. "x,y"

        :returns: True/False
        """
        if self.debug:
            log.debug("Right Click Element")
            log.debug("Location: %s", location)
        element = self.__create_element(locator)
        try:
            if location == None:
                element.Click(System.Windows.Forms.MouseButtons.Right)
                return True
            else:
                if not isinstance(location, basestring):
                    raise AssertionError("Locator must be a string")
                location = [int(x) for x in location.split(',')]
                element.Click(System.Windows.Forms.MouseButtons.Right,
                          Ranorex.Location(location[0], location[1]))
                return True
        except Exception as error:
            raise AssertionError(error)

    def rn_run_application(self, app):
        """ Runs local application.

        :param app: path to application to execute
        :returns: True/False
        """

        if self.debug:
            log.debug("Run Application %s", app)
            log.debug("Working dir: %s", os.getcwd())
        Ranorex.Host.Local.RunApplication(app)
        return True

    def rn_run_application_with_parameters(self, app, params):
        """ Runs local application with parameters.

        :param app: path to application to execute
        :param params: parameters for application
        :returns: True/False
        """

        if self.debug:
            log.debug("Run Application %s With Parameters %s", app, params)
            log.debug("Working dir: %s", os.getcwd())
        Ranorex.Host.Local.RunApplication(app, params)
        return True

    def rn_run_script(self, script_path):
        """ Runs script on remote machine and returns stdout and stderr.

        :param script_path: path to script to execute
        :returns: dictionary with "stdout" and "stderr" as keys
        """

        if self.debug:
            log.debug("Run Script %s", script_path)
            log.debug("Working dir: %s", os.getcwd())
        process = subprocess.Popen([script_path],
                                   stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE)
        output = process.communicate()
        return {'stdout':output[0], 'stderr':output[1]}

    def rn_run_script_with_parameters(self, script_path, *params):
        """ Runs script on remote machine and returns stdout and stderr.

        :param script_path: path to script to execute
        :param params: parameters for script
        :returns: dictionary with "stdout" and "stderr" as keys
        """

        params = list(params)
        if self.debug:
            log.debug("Run Script %s with params %s", script_path, params)
            log.debug("Working dir: %s", os.getcwd())
        process = subprocess.Popen([script_path] + params,
                                   stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE)
        output = process.communicate()
        return {'stdout':output[0], 'stderr':output[1]}

    def rn_scroll(self, locator, amount):
        """ Hover above selected element and scroll positive or negative
        amount of wheel turns

        :param locator: xpath pointing to desired element
        :param amount: int - amount of scrolling
        :return: None
        """

        element = self.__create_element(locator)
        mouse = Ranorex.Mouse()
        mouse.MoveTo(element.Element)
        mouse.ScrollWheel(int(amount))

    def rn_select_by_index(self, locator, index):
        """ Selects item from combobox according to index.

        :param locator: xpath selector of element
        :returns: True/False
        """

        if self.debug:
            log.debug("Select By Index %s", index)
        element = self.__create_element(locator)
        selected = element.Element.GetAttributeValue("SelectedItemIndex")
        if self.debug:
            log.debug("Selected item: %s", selected)
        diff = int(selected) - int(index)
        if self.debug:
            log.debug("Diff for keypress: %s", diff)
        if diff >= 0:
            for _ in range(0, diff):
                element.PressKeys("{up}")
        elif diff < 0:
            for _ in range(0, abs(diff)):
                element.PressKeys("{down}")
        return True

    def rn_send_keys(self, locator, key_seq):
        """ Send key combination to specified element.
        Also it gets focus before executing sequence
        seq according to :
        http://msdn.microsoft.com/en-us/library/system.windows.forms.keys.aspx

        :param locator: xpath selector of element
        :param key_seq: sequence of keys to press, i.e. {Up}{Down}{AKey}...
        :returns: True/False
        """

        if self.debug:
            log.debug("Send Keys %s", key_seq)
        Ranorex.Keyboard.PrepareFocus(locator)
        Ranorex.Keyboard.Press(key_seq)
        return True

    def rn_set_focus(self, locator):
        """ Sets focus on desired location.

        :param locator: xpath selector of element
        :returns: True/False
        """

        if self.debug:
            log.debug("Set Focus")
        element = self.__create_element(locator)
        element.Focus()
        return element.HasFocus

    def rn_take_screenshot(self, locator):
        """ Takes screenshot and return it as base64.

        :param locator: xpath selector of element
        :returns: True/False
        """
        if self.debug:
            log.debug("Take Screenshot")
        element = self.__create_element(locator)
        img = element.CaptureCompressedImage()
        return img.ToBase64String()

    def rn_uncheck(self, locator):
        """ Check if element is checked. If yes it uncheck it

        :param locator: xpath selector of element
        :returns: True/False
        """

        if self.debug:
            log.debug("Uncheck")

        element = self.__create_element(locator)
        if element.Element.GetAttributeValue('Checked'):
            if self.debug:
                log.debug("Object is checked => unchecking")
            element.Click()
            return True

    def rn_wait_for_element(self, locator, timeout):
        """ Wait for element becomes on the screen.

        :param locator: xpath selector of element
        :param timeout: timeout in milliseconds
        :returns: True/False
        """

        if self.debug:
            log.debug("Wait For Element")
            log.debug("Locator: %s", locator)
            log.debug("Timeout: %s", timeout)
        Ranorex.Validate.EnableReport = False
        if Ranorex.Validate.Exists(locator, int(timeout)) is None:
            return True
        raise AssertionError("Element %s does not exists" % locator)

    def rn_wait_for_element_attribute(self, locator, attribute,
                                   expected, timeout):
        """ Wait for element attribute becomes requested value.

        :param locator: xpath selector of element
        :returns: True/False
        """

        if self.debug:
            log.debug("Wait For Element Attribute")
            log.debug("Locator: %s", locator)
            log.debug("Attribute: %s", attribute)
            log.debug("Expected: %s", expected)
            log.debug("Timeout: %s", timeout)
        curr_time = 0
        timeout = int(timeout)/1000
        while curr_time != timeout:
            value = self.get_element_attribute(locator, attribute)
            if str(value) == str(expected):
                return True
            time.sleep(5)
            curr_time += 5
        raise AssertionError("Object at location %s could not be found"
                             % locator)

    def rn_wait_for_process_to_start(self, process_name, timeout):
        """ Waits for /timeout/ seconds for process to start.

        :param process_name: name of process to wait for
        :param timeout: timeout in milliseconds
        :returns: True/False
        """

        if self.debug:
            log.debug("Wait For Process To Start")
            log.debug("Process name: %s", process_name)
            log.debug("Timeout: %s", timeout)
        curr_time = 0
        timeout = int(timeout)/1000
        while curr_time <= timeout:
            proc = subprocess.Popen(['tasklist'], stdout=subprocess.PIPE)
            out = proc.communicate()[0]
            res = out.find(process_name) != -1 if out else False
            if res:
                return True
            else:
                curr_time += 5
                time.sleep(5)
        raise AssertionError("Process %s not found within %ss" % (process_name,
                                                                  timeout))

    def rn_kill_process(self, process_name):
        """ Kills process identified by process_name

        :param process_name: name of process to kill
        :returns: True/False
        """

        if self.debug:
            log.debug("Kill Process %s", process_name)
        res = self.check_if_process_is_running(process_name)
        if self.debug:
            log.debug("Process is running: %s", res)
        if not res:
            raise AssertionError("Process %s is not running" % process_name)
        proc = subprocess.Popen(['taskkill', '/im', process_name, '/f'],
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        out = proc.communicate()
        if 'SUCCESS' in out[0]:
            if self.debug:
                log.debug("Output of killing: %s", out)
            return True
        else:
            raise AssertionError("Process %s not terminated because of: %s" %
                                 (process_name, out))

    '''
    @Company: FileString
    @Author: minh.nguyen
    @Date: September 22, 2014
    @Description: Get text from element
    '''
    def rn_get_text(self, elementLocator=None):
        """ Get element text.

        :param locator: xpath selector of element
        :returns: element's text
        """
        return self.rn_get_element_attribute(elementLocator, 'Text')

    def rn_element_text_should_be(self, elementLocator=None, text=''):
        """ Get element text and then compare to Text

        :param locator: xpath selector of element
        :returns: True/False
        """
        element_text = self.rn_get_element_attribute(elementLocator, 'Text')
        if element_text==text:
            return True
        raise AssertionError("The element text \"%s\" is different with: \"%s\"" %
                                 (element_text, text))

    def rn_get_image(self, elementLocator=None, path_to_store_image=None):
        """ capture image by elementLocator and store it in path_to_store_image

        :param locator: xpath selector of element
        :returns: N/A
        """
        if self.debug:
            log.debug("Capture Element")
        element = self.__create_element(elementLocator)
        img = element.CaptureCompressedImage()
        img.Store(path_to_store_image)
    
    def rn_set_ranorex_time_out(self,timeout=timeOut):
        if self.debug:
            log.debug("Set ranorex time out")
        self.__timeOut = timeout


    def rn_compare_two_images_using_ranorex(self, path_to_image1, path_to_image2, same_ratio):
        """ 
        compare 2 images

        :param locator: path to image1 & 2, same ratio ex: 0.9 = 90%
        :returns: True/False
        """
        if self.debug:
            log.debug("Compare 2 images using ranorex")

        img1 = Ranorex.Imaging.Load(path_to_image1)
        img2 = Ranorex.Imaging.Load(path_to_image2)

        result = Ranorex.Imaging.Compare(img1, img2)
        print 'Two images same rate:', result
        if (result >= float(same_ratio)):
            return True
        return False

    
    def rn_check_element_is_exist(self, locator):
        """ 
        check if element is existed

        :param locator: locator
        :returns: True/False
        """
        if self.debug:
            log.debug("check element is exist")

        element_created = False
        try:
            element_type = None
            element_type = self.__return_type(locator)
            element = getattr(Ranorex, element_type)(locator)
            element_created = True
            if self.debug:
                log.debug("Element at %s", locator)
                log.debug("Application object is %s", element)
        except:
            log.debug("Element %s not found!", element_type)
            return False
        
        if element_created:
            return True
        return False

    def rn_wait_for_element_exist(self, locator, time_out_in_milisecond=5000):
        """ 
        wait for elment is exist

        This keyword return True or False, you should check the return value

        :param locator: locator
        :returns: True/False
        """
        if self.debug:
            log.debug("wait for element till timeout")

        time_out_in_milisecond = float(time_out_in_milisecond)/1000
        element_created = False
        start_time = time.time()
        while(time.time()-start_time <= time_out_in_milisecond):
            try:
                element_type = None
                element_type = self.__return_type(locator)
                element = getattr(Ranorex, element_type)(locator)
                element_created = True
                if self.debug:
                    log.debug("Element at %s", locator)
                    log.debug("Application object is %s", element)
                return True
            except:
                log.debug("Element \'%s\' not found!", locator)
        return False

    def rn_input_text_without_clear_text(self, locator, text):
        """ input texts into specified locator without clear text.

        :param locator: xpath selector of element
        :param text: text value to input into element
        :returns: True/False
        """
        if self.debug:
            log.debug("Input Text: %s", text)
        element = self.__create_element(locator)
        element.PressKeys(text)
        return True


def configure_logging():
    logging.basicConfig(
            format="%(asctime)s::[%(name)s.%(levelname)s] %(message)s",
            datefmt="%I:%M:%S %p",
            level='DEBUG')
    logging.StreamHandler(sys.__stdout__)

def main():
    # get configured logger
    logger = logging.getLogger("MAIN")

    # define arguments
    parser = ArgumentParser(prog="rxconnector", description="Remote ranorex library for robot framework")
    parser.add_argument("-i","--ip", required=False, dest="ip", default="0.0.0.0")
    parser.add_argument("-p", "--port", required=False, type=int, dest="port", default=11000)

    # parse arguments
    args = parser.parse_args()

    # run server
    try:
        server = RobotRemoteServer(RanorexLibrary(), args.ip, args.port)
    except KeyboardInterrupt, e:
        logger.info("INFO: Keyboard Iterrupt: stopping server")
        server.stop_remote_server()

if __name__ == '__main__':
    configure_logging()
    main()
    sys.exit(0)
