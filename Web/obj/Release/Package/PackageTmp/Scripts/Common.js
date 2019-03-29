function PopSelect_pop_ModalDialog(title, url, w, h) {
    var winreswidth = w;
    var winresheight = h;
    var aboutbox = "";
    var filename = "FWMODALDIALOG.ASPX?title=" + escape(title)
            + "&url=" + escape(url)
    aboutbox = showModalDialog(filename, window, "dialogWidth:" + winreswidth + "px; dialogHeight:" + winresheight + "px;unadorned:no;help:no;toolbar=no;menubar=no;location=no;status=no;scrollbars=1;resizable=1");
    //open(filename, "", "");
    return aboutbox;
}

function CopyAll(fromList, toList, isRemovefromList) {
    for (i = 0; i < fromList.options.length; i++) {
        var current = fromList.options[i];
        var text = current.text;
        var value = current.value;
        if (!ExistsItem(value, toList)) {
            var item1 = document.createElement("OPTION");
            item1.value = value;
            item1.text = text;
            toList.add(item1);
        }
        if (isRemovefromList) {
            fromList.remove(i);
        }
        i--;
    }
}
function CopyTo(fromList, toList, isRemovefromList) {
    for (i = 0; i < fromList.options.length; i++) {
        var current = fromList.options[i];
        if (current.selected) {
            var text = current.text;
            var value = current.value;
            if (!ExistsItem(value, toList)) {
                var item1 = document.createElement("OPTION");
                item1.value = value;
                item1.text = text;
                toList.add(item1);
            }
            if (isRemovefromList) {
                fromList.remove(i);
            }
            i--;
        }
    }

}
function ExistsItem(moveValue, toList) {
    var isExists = false;
    for (j = 0; j < toList.options.length; j++) {
        if (moveValue == toList.options[j].value) {
            isExists = true;
            break;
        }
    }
    return isExists;
}

//自适应窗体函数
function reinitIframe(frm) {

    var iframe = document.getElementById(frm);

    try {
        iframe.height = iframe.contentWindow.document.documentElement.scrollHeight;
    }
    catch (ex) { }

}







function GRFW_TextBox_KeyDown_Init() {

    for (var i = 0; i < GRFW_Select_TextBoxs.length; i++) {
        var txt = GRFW_TextBox_FindControl(GRFW_Select_TextBoxs[i]);
        var img = GRFW_TextBox_FindControl(GRFW_Select_Images[i]);
        if (txt != null && img != null) {
            GRFW_TextBox_KeyDown_Load(txt, img);
        }
    }

}

function GRFW_TextBox_KeyDown_Load(txt, img) {
    txt.onkeypress = new Function(GetEventString(img.onclick));

}

function GRFW_TextBox_FindControl(controlName) {
    for (var i = 0; i < document.forms.length; i++) {
        var theForm = document.forms[i];
        var theControl = theForm[controlName];
        if (theControl != null) {
            return theControl;
        }
    }
    return null;
}

function GetEventString(ev) {
    var script = "var e=window.event;if(e.keyCode != 32){event.returnValue = false;	return '';}";

    if (typeof (ev) == "function") {
        ev = ev.toString();
        ev = ev.substring(ev.indexOf("{") + 1, ev.lastIndexOf("}") - 1);
    }
    else {
        ev = "";
    }
    return script + ev + "event.returnValue = false;return '';";
}



//查询DIV折叠

function plusClick(imgID, queryFormID) {
    var objImg = document.getElementById(imgID);
    var oDiv = document.getElementById(queryFormID);
    if (oDiv.style.display == 'none') {
        objImg.className = "ui-accordion-header-icon ui-icon ui-icon-triangle-1-e";
        objImg.title = "折叠";
        oDiv.style.display = 'block';
    } else {
        objImg.className = "ui-accordion-header-icon ui-icon ui-icon-triangle-1-w";
        objImg.title = "展开";
        oDiv.style.display = 'none';
    }
}

function OpenWin(url, width, height, scrollbars) {
    var winLeft = (screen.availWidth - width) / 2;
    var winTop = (screen.availHeight - height) / 3;
    var winWidth = width;
    var winHeight = height;
    window.open(url, "null", "height=" + winHeight + ",width=" + winWidth + ",left=" + winLeft + ",top=" + winTop + ",status=yes,scrollbars=" + scrollbars);
}

function OpenWinWithWinName(url, width, height, scrollbars, winname) {
    var winLeft = (screen.availWidth - width) / 2;
    var winTop = (screen.availHeight - height) / 3;
    var winWidth = width;
    var winHeight = height;
    window.open(url, winname, "height=" + winHeight + ",width=" + winWidth + ",left=" + winLeft + ",top=" + winTop + ",status=yes,scrollbars=" + scrollbars);
}

//列表以","分割,属性以"ぁ"分割
function SelectPeople(title, w, h, InputName, InputCode) {
    var _path = "http://" + window.location.href.split("/")[2] + "/";
    if (window.location.href.toString().replace("http://", "").indexOf(":") > 0) {
        _path = "http://" + window.location.href.split("/")[2] + "/"
    }
    var url = _path + "Control/UserControl/sltPerson/SelectBase.aspx"
    var filename = "FWMODALDIALOG.ASPX?title=" + escape(title) + "&url=" + escape(url)
    var splitArry;
    var winreswidth = w; ;
    var winresheight = h;
    var aboutbox = "";
    var arrayParam = new Array(window, InputName, InputCode);
    aboutbox = showModalDialog(filename, arrayParam, "dialogWidth:" + winreswidth + "px; dialogHeight:" + winresheight + "px;unadorned:no;help:no;toolbar=no;menubar=no;location=no;status=no;scrollbars=1;resizable=1");
    if (aboutbox != undefined) {
        splitArry = aboutbox.split('|');
        return splitArry[0];
    }
    return "";
}

// 功能:统一弹出窗口的调用方法，使弹出窗口自动居于屏幕中间，并根据屏幕分辨率进行自适应
// 作者:林辉锋
// 参数说明：
//     url: 弹出窗口地址，无需进行encode
//     width: 窗口宽度px，缺省为800
//     height:窗口高度px，缺省为600
//     option:window.showModalDialog的第二个参数
// 返回：弹出页面返回值
// 创建日期: 2012.4.20
function OpenDialog(url, width, height, option) {
    url = encodeURI(url);
    if (typeof (width) == 'undefined') { width = 800; }
    if (typeof (height) == 'undefined') { height = 600; }
    if (typeof (option) == 'undefined') { option = ''; }
    var left = (screen.width - width) / 2;
    var top = (screen.height - height) / 2;
    return window.showModalDialog(url, option,
        "scroll:auto;dialogLeft:" + left
        + "px;dialogTop:" + top
        + "px;dialogHeight:" + height
        + "px;dialogWidth:" + width + "px;center:1;status:0");
}

/* 发起Ajax请求 */
// url：发送请求的地址
// data：发送到服务器的数据，json格式
// dataType：服务器返回的数据类型，如：xml、text、html、json、script等
// sucess：请求成功后的回调函数
// async：是否异步请求(只有明确指定为false时才为同步)
// type：请求类型，默认GET
function ajaxRequest(url, data, dataType, sucess, async, type) {

    $.ajax(
        {
            type: (type === "POST" ? "POST" : "GET"),
            url: url,
            async: !(async === false),
            data: data,
            dataType: dataType,
            success: sucess,
            beforeSend: ajaxLoading,
            complete: ajaxComplete,
            error: ajaxError
        });
}

/* Ajax请求中 */
function ajaxLoading(xmlHttpRequest) {

}

/* Ajax请求完成 */
function ajaxComplete(xmlHttpRequest, textStatus) {

}

/* Ajax请求失败时给出提示 */
function ajaxError(xmlHttpRequest, textStatus, errorThrown) {
    alert("数据获取或操作失败。\n\n代码：" + xmlHttpRequest.status + "\n信息：" + xmlHttpRequest.statusText);
}


/* 弹出一个页面，此页面无工具栏 */
function PopWindow(link) {
    var top = (window.screen.availHeight - 330) / 2;
    var left = (window.screen.availWidth - 610) / 2;
    var newWindow = window.open(link, "_blank", 'resizable=yes,location=no,menubar=no,status=no,scrollbars=yes,height=300,width=600,top=' + top + ',left=' + left);
    return false;
}

//获取表格某一行的第几个标签名为tagName的控件
function getObjTR(table, rowIndex, tagName, tagIndex) {
    var obj = null;
    if (table.rows(rowIndex).getElementsByTagName(tagName).length > tagIndex) {
        obj = table.rows(rowIndex).getElementsByTagName(tagName).item(tagIndex);
    }
    return obj;
}

//获取表格某一行某列的第几个标签名为tagName的控件
function getObjTRTD(table, rowIndex, cellIndex, tagName, tagIndex) {
    var obj = null;
    if (table.rows(rowIndex).cells(cellIndex).getElementsByTagName(tagName).length > tagIndex) {
        obj = table.rows(rowIndex).cells(cellIndex).getElementsByTagName(tagName).item(tagIndex);
    }
    return obj;
}


function SetTreeHeight(id) {
    var height = parent.parent.window.screen.availHeight;
    var tree = document.getElementById(id);
    $(tree).css("height", height - 180);
}

//查询条件折叠或显示
function more() {
    if (document.getElementById('super_search').style.display == '') {
        document.getElementById('super_search').style.display = 'none';
        document.getElementById('btnMore').className = 'ui-btn ui-btn-view';
        document.getElementById('btnMore').value = '更多筛选条件';
    }
    else {
        document.getElementById('super_search').style.display = '';
        document.getElementById('btnMore').className = 'ui-btn ui-btn-hidden ';
        document.getElementById('btnMore').value = '精简筛选条件';
    }

}

// 获取页面的Html控件
function getObj(id) {
    if (!id) {
        return null;
    }
    return document.getElementById(id);
}

// 获取页面数据区域容器元素
function getDataAreaObj() {
    return getObj('divMPList');
}

// 获取页面的Html控件数组
function getObjs(name) {
    return document.getElementsByName(name);
}

// 获取模式窗口父页面的Html控件
function getObjD(id) {
    if (!id) {
        return null;
    }
    return window.dialogArguments.document.getElementById(id);
}

// 获取一般弹出(非模式)窗口父页面的Html控件
function getObjO(id) {
    if (!id) {
        return null;
    }
    return window.opener.document.getElementById(id);
}

// 获取父页面的Html控件
function getObjP(id) {
    if (!id) {
        return null;
    }
    return window.parent.document.getElementById(id);
}

// 获取祖页面的Html控件
function getObjPP(id) {
    if (!id) {
        return null;
    }
    return window.parent.parent.document.getElementById(id);
}

// 获取本页某个框架(或内嵌框架)里的Html控件
function getObjF(frameName, id) {
    if (!frameName || !id) {
        return null;
    }
    return window.frames(frameName).document.getElementById(id);
}

// 获取父页面某个框架(或内嵌框架)里的Html控件
function getObjPF(frameName, id) {
    if (!frameName || !id) {
        return null;
    }
    return window.parent.frames(frameName).document.getElementById(id);
}

// 获取顶页面的Html控件
function getObjT(id) {
    if (!id) {
        return null;
    }
    return window.top.document.getElementById(id);
}

// 获取容器container中标记为tagName的HTML控件数组中的第tagIndex个控件
// 容器可以是table、tr、td、div等可以含子元素的元素
function getObjC(container, tagName, tagIndex) {
    var obj = null;
    if (container.getElementsByTagName(tagName).length > tagIndex) {
        obj = container.getElementsByTagName(tagName).item(tagIndex);
    }
    return obj;
}

// 获取表格table中第rowIndex行的标记为tagName的HTML控件数组中的第tagIndex个控件
function getObjTR(table, rowIndex, tagName, tagIndex) {
    var obj = null;
    if (table.rows(rowIndex).getElementsByTagName(tagName).length > tagIndex) {
        obj = table.rows(rowIndex).getElementsByTagName(tagName).item(tagIndex);
    }
    return obj;
}

// 获取表格table中第rowIndex行第colIndex列的标记为tagName的HTML控件数组中的第tagIndex个控件
function getObjTC(table, rowIndex, colIndex, tagName, tagIndex) {
    var obj = null;
    if (table.rows[rowIndex].cells[colIndex].getElementsByTagName(tagName).length > tagIndex) {
        obj = table.rows[rowIndex].cells[colIndex].getElementsByTagName(tagName).item(tagIndex);
    }
    return obj;
}

//获取表格某一行某列的第几个标签名为tagName的控件
function getObjTRTD(table, rowIndex, cellIndex, tagName, tagIndex) {
    var obj = null;
    if (table.rows(rowIndex).cells(cellIndex).getElementsByTagName(tagName).length > tagIndex) {
        obj = table.rows(rowIndex).cells(cellIndex).getElementsByTagName(tagName).item(tagIndex);
    }
    return obj;
}

//弹出窗口居中显示
function openwindow(url, name, iWidth, iHeight) {
    var iTop = (window.screen.availHeight - 30 - iHeight) / 2; //获得窗口的垂直位置;
    var iLeft = (window.screen.availWidth - 10 - iWidth) / 2; //获得窗口的水平位置;
    window.open(url, name, 'height=' + iHeight + ',,innerHeight=' + iHeight + ',width=' + iWidth + ',innerWidth=' + iWidth + ',top=' + iTop + ',left=' + iLeft + ',toolbar=no,menubar=no,scrollbars=yes,resizeable=no,location=no,status=no');
}


function OpenWindow(link) {
    var newWindow = window.open(link, "_blank", "left=0,top=0,width=" + (screen.availWidth - 10) + ",height=" + (screen.availHeight - 50) + "toolbar=no, menubar=no, scrollbars=yes, resizable=yes,location=no, status=no", "");
    if (document.all) {
        newWindow.moveTo(0, 0);
        newWindow.resizeTo(screen.width, screen.availHeight);
    }
}


function OpenWindowInner(link, currtPage) {
    if (currtPage == "Y") {
        this.location.href = link;
    }
    else {
        this.opener.location.href = link;
    }
}



//全选,反选
function selALL(id) {
    var tb = document.getElementById(id);
    if (tb) {
        if (getObjTR(tb, 0, 'input', 0).checked) {
            for (var i = 0; i < tb.rows.length; i++) { getObjTR(tb, i, 'input', 0).checked = true; }
        }
        else {
            for (var i = 0; i < tb.rows.length; i++) { getObjTR(tb, i, 'input', 0).checked = false; }
        }
    }
}

//判断是否选择
function CheckSelect(id) {
    var tb = document.getElementById(id);
    var count = 0;
    if (tb) {
        for (var i = 0; i < tb.rows.length; i++) {
            if (getObjTR(tb, i, 'input', 0) != undefined && getObjTR(tb, i, 'input', 0).checked) {
                count++;
            }
        }
    }
    if (count == 0) alert('请选择！');
    return count > 0;
}

//判断是否选择
function CheckSelectOne(id) {
    var tb = document.getElementById(id);
    var count = 0;
    if (tb) {
        for (var i = 0; i < tb.rows.length; i++) {
            if (getObjTR(tb, i, 'input', 0) != undefined && getObjTR(tb, i, 'input', 0).checked) {
                count++;
            }
        }
    }
    if (count == 0 || count > 1) alert('有且只能选择一项！');
    return count == 1;
}

