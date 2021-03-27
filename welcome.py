from kivy.config import Config
Config.set('graphics', 'resizable', 0)
Config.set('kivy', 'window_icon', 'hrm_logo.png')

from kivy.app import App
from kivy.core.window import Window
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.lang import Builder
from kivy.properties import ObjectProperty
from kivy.clock import Clock
import requests
from zk import ZK, const
from xlwt import Workbook
import pandas as pd



class FirstWindow(Screen):

    check = ObjectProperty(None)

    def next(self):
        try:
            check = self.ids.check
            data2 = pd.read_excel('second.xls')
            d2 = list(data2)


            payload = {"purchase_key": d2[1],
                       "product_key": "20386502",
                       "domain": d2[0]
                       }

            API = "https://store.bdtask.com/class.api.php"
            r = requests.get(url=API, params=payload)
            print(r.text)
            data = r.json()

            if data['status'] == True:
                check.text = "Next"
                print("True")
            else:
                check.text = "Verify"
                print("False1")

        except Exception:
                check.text = "Verify"
                print("False2")



class SecondWindow(Screen):

    url = ObjectProperty(None)
    key = ObjectProperty(None)
    info = ObjectProperty(None)


    def next(self):

        try:
            url = str(self.url.text)
            key = self.key.text
            info = self.ids.info

            if (url and key) != "":
                wb = Workbook()
                sheet1 = wb.add_sheet('Sheet 1')
                sheet1.write(0, 0, url)
                sheet1.write(0, 1, key)
                wb.save('second.xls')

            else:
                print("Fill the requirement")


            data2 = pd.read_excel('second.xls')
            d2 = list(data2)

            payload = {"purchase_key": d2[1],
                       "product_key": "20386502",
                       "domain": d2[0]
                       }

            API = "https://store.bdtask.com/class.api.php"
            r = requests.get(url=API, params=payload)
            print(r.text)
            data = r.json()

            if data['status'] == True:
                info.text = "Next"

            else:
                info.text = "Verify"

        except Exception:
                info.text = "Verify"
                print("False2")



class ThirdWindow(Screen):

    ip = ObjectProperty(None)
    port = ObjectProperty(None)

    def save_data(self):
        ip = self.ip.text
        port = self.port.text
        sinfo = self.ids.sinfo

        if (ip and port) != "":
            wb = Workbook()
            sheet1 = wb.add_sheet('Sheet 1')
            sheet1.write(0, 0, ip)
            sheet1.write(0, 1, port)
            wb.save('third.xls')
            sinfo.text = "[color=#056608]All information save successfully[/color]"
        else:
            sinfo.text ="[color=#FF0000]Fill all the required field[/color]"

        self.ip.text = ""
        self.port.text = ""



class SubMenuHRMApp(Screen):

    search = ObjectProperty(None)

    def hrm_app(self):
        search = self.search.text
        einfo = self.ids.einfo


        try:
            payload = {"employee_id": search}
            API = "https://adminhr.bdtask.com/api_handler/get_employee_by_id"
            r = requests.post(API, data=payload)
            data = r.json()

            if data["status"] == "ok":
                data_list = pd.read_excel('third.xls')
                d = list(data_list)
                conn = None
                zk = ZK(d[0], port=int(d[1]), timeout=5, password=0, force_udp=False, ommit_ping=False)
                print(type(d[1]))
                conn = zk.connect()
                conn.disable_device()
                users = conn.get_users()
                user_id_list = []
                for user in users:
                    user_id_list.append(user.user_id)

                if (search in user_id_list):
                    einfo.text = "[color=#056608]Id already exists on Server and ZKT[/color]"

                else:
                    name = data["data"]["first_name"] + " " + data["data"]["last_name"]
                    id = data["data"]["employee_id"]
                    conn.set_user(uid=None, name=name, privilege=0, password='', group_id='', user_id=id, card=0)
                    einfo.text = "[color=#056608]Welcome! Id has been created on ZKT[/color]"

                conn.test_voice()
                conn.enable_device()
                conn.disconnect()
            else:
                einfo.text = "[color=#FF0000]This Employee info isn't exist on Server and ZKT[/color]"

        except Exception as e:
            einfo.text = "Error"


    def delete_employee(self):

        search = self.search.text
        einfo = self.ids.einfo
        try:
            data = pd.read_excel('third.xls')
            d = list(data)
            conn = None
            zk = ZK(d[0], port=int(d[1]), timeout=5, password=0, force_udp=False, ommit_ping=False)
            conn = zk.connect()

            conn.disable_device()
            conn.delete_user(user_id=search)

            conn.test_voice()
            conn.enable_device()
            conn.disconnect()
            einfo.text = "Delete successfully"

        except Exception as e:
            einfo.text = "Delete Error : {}".format(e)



class WindowManager(ScreenManager):
    pass


kv = Builder.load_file('hrm.kv')



class MyApp(App):


    def build(self):

        Window.size = (414, 736)
        Window.clearcolor = (1, 1, 1, 1)
        self.title = 'HRM APP'
        Clock.schedule_interval(self.press, 3600)
        return kv

    def press(self, *args):

        data2 = pd.read_excel('second.xls')
        d2 = list(data2)
        data = pd.read_excel('third.xls')
        d = list(data)

        conn = None
        zk = ZK(d[0], port=int(d[1]), timeout=5, password=0, force_udp=False, ommit_ping=False)
        try:
            payload = {"status": 1}
            API = d2[0] + "/api_handler/check_status"
            r = requests.post(API, data=payload)
            testdata = r.json()
            if testdata['status'] == "ok":
                try:
                    conn = zk.connect()
                    conn.disable_device()
                    attendances = conn.get_attendance()
                    for index, attendance in enumerate(attendances):
                        payload = {
                            "uid": attendance.user_id,
                            "id": attendance.uid,
                            "state": attendance.status,
                            "time": attendance.timestamp
                        }
                        API = d2[0] + "/api_handler/create_attendence"
                        r = requests.post(API, data=payload)
                        print("Hello", r.text)
                    conn.clear_attendance()
                    conn.enable_device()

                except Exception as e:
                    print("Process terminate : {}".format(e))

                finally:
                    if conn:
                        conn.disconnect()
            else:
                print("API is not working properly")

        except Exception as e:
            print("Process terminate : {}. Check your internet connection".format(e))


try:
    MyApp().run()
except Exception as e:
    with open('error.txt', 'a+') as f:
        f.write(str(e) + '\n')
