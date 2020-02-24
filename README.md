# WITClassScheduleToCalendar
武汉工程大学教务处（强智）课程表转.ics格式，Python编写，用于将课程表导入iOS、macOS平台自带的Calendar应用，以及通用iCal格式的Google Calendar和Windows自带的Calendar软件。

在教务系统网站[武汉工程大学教务系统](http://jwxt.wit.edu.cn/jsxsd)登陆
![教务系统登陆](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/jwxt3.png)
![教务系统课表入口](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/jwxt.png)
![教务系统打印](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/jwxt2.png)
点击```打印```即可下载课表的.xls文件

安装依赖：
```pip install ics```
```pip install xlrd```
```pip install datetime```

自行修改
```filepath```
```output_filepath```兩個變量

運行```main.py```:
![输入index和第一周星期一的日期](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/sheet_index.png)
輸入課程表在excel文件中的sheet index  
輸入本學期第一週的星期一的日期

得到.ics文件

附Google Calendar导入方法
![设置](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/google_calendar.png)
![设置2](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/google_calendar2.png)
![设置3](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/google_calendar3.png)
![设置4](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/google_calendar4.png)
macOS自带Calendar导入iCloud方法
![macOS](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/macOS.png)
双击打开即可导入
![macOS2](https://github.com/DOROMOLLL/WITClassScheduleToCalendar/blob/master/images/macOS2.png)
