<!--#include file="md5.asp"-->
<!--#include file="jsonObject-class.asp"-->
<!--#include file="util.asp"-->

<%
' 阿里大鱼APP类
Class AlidayuApp
	Dim url
	Dim appKey
	Dim appSecret
	Dim defaultConfig

	Private Sub Class_Initialize
		Response.Write("<br><br>---- Class_Initialize <br>")

		url = "http://gw.api.taobao.com/router/rest"
		Set defaultConfig = New JSONobject

		defaultConfig.Add "method", ""
		defaultConfig.Add "format", "json"
		defaultConfig.Add "v", "2.0"
		defaultConfig.Add "sign_method", "md5"	
	End Sub

	Private Sub Class_Terminate
		Response.Write("<br><br>---- Class_Terminate <br>")
		Set url = Nothing
	    Set appKey = Nothing
		Set appSecret = Nothing
		Set defaultConfig = Nothing
	End Sub

	' 配置app key & secret
	Sub config(mKey, mSecret)
		Response.Write("<br><br>---- config: " & mKey & ", " & mSecret & " <br>")

		appKey = mKey
		appSecret = mSecret
		defaultConfig.Add "app_key", mKey
	End Sub

	' 签名
	' @param {object} params 参数
	Function getSign(params)
		Response.Write("<br><br>---- getSign <br>")

		Dim i, mSign, outParams, keyArray, mStr, outArr()

		Set outParams = New JSONobject

		outParams.merge(defaultConfig)
		outParams.merge(params)
		outParams.change "timestamp", formatDate(Now)

		Response.write("outParams.Serialize(): " & outParams.Serialize() & "<br>")

		keyArray = Split(outParams.keys(), ",") 
		keyArray = sortArr(keyArray)

		ReDim outArr(UBound(keyArray) + 3)
		outArr(0) = appSecret

		for i = 0 To UBound(keyArray)
			Key = keyArray(i)
			outArr(i + 1) = Key & outParams(key)
		next 

		outArr(i + 2) = appSecret
		mStr = Join(outArr, "")
		mSign = UCase(asp_md5(mStr))

		Response.write("sort str: " & mStr & "<br>")
		Response.write("sign: " & mSign & "<br>")

		outParams.change "sign", mSign

		getSign = outParams.Serialize()
	End Function 

	' 请求mtop服务
	' @param {object} params 参数
	' @param {funciton} callback 回调
	Function request(params)
		Response.Write("<br><br>---- request <br>")

		Dim ServerXmlHttp, paramsStr, signStr

		signStr = Me.getSign(params)
		params.Parse(signStr)
		Response.write("parse signStr success: " & signStr & "<br>")

		params.change "sms_free_sign_name", Server.URLEncode(params("sms_free_sign_name"))
		params.change "timestamp", Server.URLEncode(params("timestamp"))
		
		paramsStr = params.params()

		Response.write("paramsStr: " & paramsStr & "<br>")

		Set ServerXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
		ServerXmlHttp.open "POST", url
		ServerXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=utf-8"
		ServerXmlHttp.setRequestHeader "app_key", appKey
		ServerXmlHttp.setRequestHeader "Content-Length", Len(paramsStr)
		ServerXmlHttp.send paramsStr

		If ServerXmlHttp.status = 200 Then
			TextResponse = ServerXmlHttp.responseText
			Response.write("request success: " & TextResponse & "<br>")
		Else
			' Handle missing response or other errors here
			Response.write("request error: " & url & "<br>")
		End If

		Set ServerXmlHttp = Nothing

		request = TextResponse
	End Function 


 ' 短信发送
 ' @param {object} params 参数
 ' - params
 '    @param {string} sms_type 短信类型 normal
 '    @param {string} sms_free_sign_name 短信签名 传入的短信签名必须是在阿里大鱼“管理中心-短信签名管理”中的可用签名。
 '    @param {json} sms_param 短信模板变量 传参规则{"key":"value"}，key的名字须和申请模板中的变量名一致，多个变量之间以逗号隔开
 '    @param {string} rec_num 短信接收号码 支持单个或多个手机号码，传入号码为11位手机号码，不能加0或+86。群发短信需传入多个号码，以英文逗号分隔，一次调用最多传入200个号码
 '    @param {string} sms_template_code 短信模板ID 传入的模板必须是在阿里大鱼“管理中心-短信模板管理”中的可用模板。示例：SMS_585014
	Sub smsSend(params)
		Response.Write("<br><br>---- smsSend <br>")

		Dim defaultParams
		Dim reqParams

		Set defaultParams = New JSONobject
		Set reqParams = New JSONobject

		defaultParams.add "method", "alibaba.aliqin.fc.sms.num.send"
		defaultParams.add "sms_type", "normal"

		reqParams.merge(defaultParams)
		reqParams.merge(params)
		
		Me.request(reqParams)
	End Sub 

	' 短信发送记录查询
	' @param {object} params 参数
	' - params
	'	 @param {string} rec_num 短信接收号码
	'	 @param {string} query_date 短信发送日期，支持近30天记录查询，格式yyyyMMdd
	'	 @param {number} current_page 分页参数,页码
	'	 @param {number} page_size 分页参数，每页数量。最大值100
	Sub smsQuery(params)
		Response.Write("<br><br>---- smsQuery <br>")

		Dim defaultParams
		Dim reqParams

		Set defaultParams = New JSONobject
		Set reqParams = New JSONobject

		defaultParams.add "method", "alibaba.aliqin.fc.sms.num.query"

		reqParams.merge(defaultParams)
		reqParams.merge(params)
		
		Me.request(reqParams)
	End Sub 

	' 语音通知
	' @param {object} params 参数
	' - params
	'	 @param {string} called_num 被叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
	'	 @param {string} called_show_num 被叫号显，传入的显示号码必须是阿里大鱼“管理中心-号码管理”中申请通过的号码
	'	 @param {string} voice_code 语音文件ID，传入的语音文件必须是在阿里大鱼“管理中心-语音文件管理”中的可用语音文件
	Sub voiceSinglecall (params)
		Response.Write("<br><br>---- smsQuery <br>")

		Dim defaultParams
		Dim reqParams

		Set defaultParams = New JSONobject
		Set reqParams = New JSONobject

		defaultParams.add "method", "alibaba.aliqin.fc.voice.num.singlecall"

		reqParams.merge(defaultParams)
		reqParams.merge(params)
		
		Me.request(reqParams)
	End Sub 

	' 文本转语音通知
	' @param {object} params 参数
	' - params
	'	 @param {json} tts_param 文本转语音（TTS）模板变量，传参规则{"key"："value"}，key的名字须和TTS模板中的变量名一致，多个变量之间以逗号隔开
	'	 @param {string} called_num 被叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
	'	 @param {string} called_show_num 被叫号显，传入的显示号码必须是阿里大鱼“管理中心-号码管理”中申请或购买的号码
	'	 @param {string} tts_code TTS模板ID，传入的模板必须是在阿里大鱼“管理中心-语音TTS模板管理”中的可用模板
	Sub ttsSinglecall (params)
		Response.Write("<br><br>---- smsQuery <br>")

		Dim defaultParams
		Dim reqParams

		Set defaultParams = New JSONobject
		Set reqParams = New JSONobject

		defaultParams.add "method", "alibaba.aliqin.fc.tts.num.singlecall"

		reqParams.merge(defaultParams)
		reqParams.merge(params)
		
		Me.request(reqParams)
	End Sub 

	' 语音双呼
	' @param {object} params 参数
	' - params
	'	 @param {string} caller_num 主叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
	'	 @param {string} caller_show_num 主叫号码侧的号码显示，传入的显示号码必须是阿里大鱼“管理中心-号码管理”中申请通过的号码。显示号码格式如下057188773344，4001112222，95500
	'	 @param {string} called_num 被叫号码，支持国内手机号与固话号码,格式如下057188773344,13911112222,4001112222,95500
	'	 @param {string} called_show_num 被叫号码侧的号码显示，传入的显示号码可以是阿里大鱼“管理中心-号码管理”中申请通过的号码。显示号码格式如下057188773344，4001112222，95500。显示号码也可以为主叫号码。
	Sub voiceDoublecall (params)
		Response.Write("<br><br>---- smsQuery <br>")

		Dim defaultParams
		Dim reqParams

		Set defaultParams = New JSONobject
		Set reqParams = New JSONobject

		defaultParams.add "method", "alibaba.aliqin.fc.sms.num.query"

		reqParams.merge(defaultParams)
		reqParams.merge(params)
		
		Me.request(reqParams)
	End Sub 

End Class 
%>