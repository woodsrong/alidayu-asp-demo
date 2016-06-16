<!--#include file="alidayu-sdk/index.asp"-->

<%
'调用阿里大鱼SDK发送短信示例

Response.charset = "utf-8"
Response.Write("start <br>")

Dim mApp
Dim mParams

Set mApp = New AlidayuApp
Set mParams = New JSONobject

'修改为自己的app key和app secret，以下为测试key和secret，已失效
mApp.config "23275748", "bdfc910d42a7b421400c4b1c9c2d65c9"

'配置短信签名
mParams.add "sms_free_sign_name", "大鱼测试"
'配置接受短信手机号码
mParams.add "rec_num", "18688705716"
'配置短信模板
mParams.add "sms_template_code", "SMS_10370738"

mApp.smsSend(mParams)
%>