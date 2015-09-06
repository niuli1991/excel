<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ taglib prefix="s" uri="/struts-tags"%>
<%
        String jsonDatas = (String)request.getAttribute("jsonDatas");
        String colName = (String)request.getAttribute("colsName");
        String path = request.getContextPath();
		String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    <base href="<%=basePath%>">
    
    <title>My JSP 'excelUpLoad.jsp' starting page</title>
    
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache">
	<meta http-equiv="expires" content="0">    
	<meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
	<meta http-equiv="description" content="This is my page">
	<!--
	<link rel="stylesheet" type="text/css" href="styles.css">
	-->
	<script type="text/javascript" src="<%=basePath%>/javascript/jquery.min.js"></script>
  </head>
  <script>
     // var jsonDatas = <%=jsonDatas%> ;
      var maxsize = 2 * 1024 * 1024;//2M
      var errMsg = "上传的附件文件不能超过2MB！！！";
      var tipMsg = "您的浏览器暂不支持计算上传文件的大小，确保上传文件不要超过2M，建议使用IE、FireFox、Chrome浏览器。";
      var browserCfg = {
      };
      var ua = window.navigator.userAgent;
      if (ua.indexOf("MSIE") >= 1) {
          browserCfg.ie = true;
      }
      else if (ua.indexOf("Firefox") >= 1) {
          browserCfg.firefox = true;
      }
      else if (ua.indexOf("Chrome") >= 1) {
          browserCfg.chrome = true;
      }
      //上传文件的校验
      function check() {
          //文件空校验
          var flage = false;
          if ($("input[name='file']").val().length) {
              flage = true;
          }
          if (!flage) {
              alert('无上传文件,请添加文件！');
              return false;
          }
          //文件类型校验
          $obj = $("input[name='file']");
          var filename = $obj.val().lastIndexOf('\\') !=  - 1 ? $obj.val().substring($obj.val().lastIndexOf('\\') + 1) : $obj.val();
          var fileType = filename.split(".")[filename.split(".").length - 1];
          if (fileType != "xls") {
              alert("文件格式不对，请重新上传Excel文件！");
              return false;
          }
          flage = checkSize();

          if (!flage) {
              return true;
          }
      }
      //文件大小校验，最大2MB
      function checkSize() {
          try {
              var obj_file = $("input[name='file']")[0];
              var filesize = 0;
              if (browserCfg.firefox || browserCfg.chrome) {
                  filesize = obj_file.files[0].size;
              }
              else if (browserCfg.ie) {
                  var obj_file_2 = $("input[name='file']")[0];
                  obj_file_2.dynsrc = obj_file.value;
                  //                  alert(obj_file);
                  //filesize = obj_file_2.files[0].size;
              }
              else {
                  alert(tipMsg);
                  return true;
              }
              if (filesize ==  - 1) {
                  alert(tipMsg);
                  return true;
              }
              else if (filesize > maxsize) {
                  alert(errMsg);
                  return false;
              }
          }
          catch (e) {
              alert(e);
              return false;
          }
      }

      function loadData() {
          var jsonDatas = $("#jsonDatas").val(); 
          if (jsonDatas == "") {
              return false;
          }
          $("#data").val(jsonDatas);
          jsonDatas = eval('(' + jsonDatas + ')');
          var row = jsonDatas.row;
          var colName = $("#colName").val();
          
      	  $("#cols").val(colName);
          colName = colName.split("_");
          $("#datas").attr("style", "border:1px solid #95b8e7");
          for (var item = 0;item < row.length;item++) {
              var td = "";
              for (var index = 0;index < colName.length;index++) {
                  if (item == 0) {
                      td = td + '<td class="tdleft" style="line-height:15px"><div class="textdiv" style="text-align:center">' + row[item][colName[index]] + '<\/div><\/td>';
                  }
                  else if (typeof (row[item][colName[index]]) != "undefined") {
                      if (colName[index] == "result") {
                          if (row[item][colName[index]] == "成功") {
                              td = td + '<td class="td_end" style="line-height:15px"><div class="textdiv" style="color:green;text-align:center">' + row[item][colName[index]] + '<\/div><\/td>';
                          }
                          else {
                              td = td + '<td class="td_end" style="line-height:15px"><div class="textdiv" style="color:red;text-align:center">' + row[item][colName[index]] + '<\/div><\/td>';
                          }
                      }
                      else {
                          td = td + '<td class="tdright" style="line-height:15px"><div class="textdiv" style="text-align:center">' + row[item][colName[index]] + '<\/div><\/td>';
                      }
                  }
                  else {
                      td = td + '<td class="tdright" style="line-height:15px"><div class="textdiv" style="text-align:center"><\/div><\/td>';
                  }

              }
              $("#datas").append('<tr>' + td + '<\/td>');
          }
      }
      
      function downloadCheck(){
      	if($("#jsonDatas").val() == ""){
      		alert("请先上传excel！");
      		return false;
      	}
      }
      $(function () {
		loadData();
      })
    </script>
    <style type="text/css">
        input{ vertical-align:middle; margin:0; padding:0}
        .file-box{ position:relative;width:340px}
        .txt{ height:22px; border:1px solid #cdcdcd; width:180px;cursor:pointer;}
        .btn{ background-color:#FFF; border:1px solid #CDCDCD;height:24px; width:70px;cursor:pointer;}
        .file{ position:absolute; top:0; right:80px; height:24px; filter:alpha(opacity:0);opacity: 0;width:260px;cursor:pointer; }
    </style>
  <body>
    <div class="file-box">
            <s:form action="excelUpLoad" method="post"
                    namespace="/" enctype="multipart/form-data"
                    theme="simple" id="file" onsubmit="return check()">
                <s:fielderror name="file"></s:fielderror>
                <s:actionerror/>
                <table cellpadding="0" cellspacing="0" width="360px">
                    <tr>
                        <td>
                            <div style="padding:3px;">
                                <input type='text' name='textfield'
                                       id='textfield' class='txt'/>
                                 
                                <input type='button' class='btn' value='浏览...'/>
                                 
                                <input type="file" style="margin-top:12px"
                                       name="file" class="file" id="fileField"
                                       size="28"
                                       onchange="document.getElementById('textfield').value=this.value"/>
                            </div>
                        </td>
                        <td>
                            <div style="padding:3px;">                                 
                                <input type="submit" class="btn" value="上传文件"/>
                                <input type="hidden" id="jsonDatas" name="jsonDatas" value='${jsonDatas}'/>
                                <input type="hidden" id="colName" name="colName" value='${colName}'/>
                            </div>
                        </td>             
                    </tr>
                </table>
            </s:form>
        </div>
        <s:form action="excelDownLoad" method="post"namespace="/" enctype="multipart/form-data" theme="simple" id="" onsubmit="return downloadCheck();">
        	<input type="hidden" id="data" name="data" value=''>
        	<input type="hidden" id="cols" name="cols" value=''>
        	<input type="submit" class="btn" value="下载excel"/>
        </s:form>
        <div style="padding:10px">
            <table cellpadding="1" cellspacing="1" id="datas"></table>
        </div>
  </body>
</html>
