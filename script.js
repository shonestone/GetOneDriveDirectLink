const curMaskVersion = 4; // 当前掩码设置版本

// 仿 GM_* API，用 localStorage 实现
if (typeof(GM_getValue) == "undefined") {
    var GM_getValue = function(name, type){
        var value = localStorage.getItem(name);
        if (value == undefined) return value;
        if ((/^(?:true|false)$/i.test(value) && type == undefined) || type == "boolean") {
            return /^true$/i.test(value);
        } else if((/^\-?[\d\.]+$/i.test(value) && type == undefined) || type == "number") {
            return Number(value);
        } else {
            return value;
        }
    }
}
if (typeof(GM_setValue) == "undefined") {
    var GM_setValue = function(name, value){
        localStorage.setItem(name, value);
    }
}
if (typeof(GM_deleteValue) == "undefined") {
    var GM_deleteValue = function(name){
        localStorage.removeItem(name);
    }
}
if (typeof(GM_listValues) == "undefined") {
    var GM_listValues = function(){
        var keys = [];
        for (var ki=0; ki<localStorage.length; ki++) {
            keys.push(localStorage.key(ki));
        }
        return keys;
    }
}

// 掩码对象
var maskObj = function(name,content) {
    this.name = name;
    this.content = content;
    return this;
};
var masks = [];
var mask_list = null;
var mask_name = null;
var mask_content = null;
var outinfo = null;
var outcontent = null;

function addNewMask(name,content) {
    var mask = new maskObj(name,content);
    masks.push(mask);
    var opt = new Option(name + " : " + content, content);
    mask_list.options.add(opt);
}
function save_mask_local() {
    var maskstr = JSON.stringify(masks);
    GM_setValue("godl-masks",maskstr);
    GM_setValue("godl-mask-index",mask_list.selectedIndex);
}
function load_mask_local() {
    var maskstr = GM_getValue("godl-masks");
    var masksCfg;
    try { masksCfg = JSON.parse(maskstr); } catch (e) { masksCfg = null; }
    
    if (!Array.isArray(masksCfg) ||
        ((parseInt(GM_getValue("new-mask-version"),10) || 1)<curMaskVersion)
    ) {
        addNewMask("长链接 (webUrl)", "${file.name} => ${file.webUrl}");
        addNewMask("直链 (downloadUrl)", "${file.name} => ${file['@microsoft.graph.downloadUrl']}");
        addNewMask("短链 (createLink)", "${file.name} => %{awaitCreateShortLink(file.id)}");

        GM_setValue("new-mask-version",curMaskVersion);
    } else {
        masksCfg.forEach(function(item){
            addNewMask(item.name,item.content);
        });
    }

    var mask_index = parseInt(GM_getValue("godl-mask-index"),10) || 0;
    mask_list.selectedIndex = mask_index;
}

// 输出处理
function do_error(e) {
    outinfo.innerHTML = "发生错误";
    outcontent.value = e.toString();
}
function do_cancel() {
    outinfo.innerHTML = "取消操作";
}
async function do_success(files) {
    redata = files;
    console.log("返回文件数据:", redata);

    // 保存 access_token
    window._accessToken = files.accessToken;

    await generate_output(redata);
}

async function generate_output(files) {
    var mask = masks[mask_list.selectedIndex] || masks[0];
    var filearr = files.value;
    
    outinfo.innerHTML = "共选择 " + filearr.length + " 个文件。";
    if (filearr.some(function(item){
        return item.shared == undefined || item.shared.scope != "anonymous";
    })){
        outinfo.innerHTML += "存在非公共权限文件，注意添加通行许可代码。";
    }

    const outStrArr = [];
    for (let i=0; i<filearr.length; i++) {
        let outStr = await showMask(mask.content, filearr[i], i);
        outStrArr.push(outStr);
    }
    outcontent.value = outStrArr.join("\n");
}

// 掩码替换
async function showMask(str,file,index) {
    // 支持异步调用
    let newTxt = eval("`" + str +"`");

    // 处理 %{...} 表达式
    const pattern = "%{([^}]+)}";
    let rs = null;
    while (( rs = new RegExp(pattern).exec(newTxt) ) != null) {
        let mskO = rs[0], mskN = rs[1];
        if (mskN != undefined) {
            try {
                let evTemp = await eval(mskN);
                if (evTemp!=undefined)
                    newTxt = newTxt.replace(mskO, evTemp.toString());
                else
                    newTxt = newTxt.replace(mskO, "");
            } catch(e) {
                console.error("掩码异常:", e);
            }
        }
    }
    return newTxt;
}

// 调用 Graph API 生成短链
async function awaitCreateShortLink(itemId) {
    if (!window._accessToken) return "无token";
    const url = "https://graph.microsoft.com/v1.0/me/drive/items/" + itemId + "/createLink";
    const body = { type: "view", scope: "anonymous" };

    const resp = await fetch(url, {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + window._accessToken,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    });

    if (!resp.ok) {
        return "Graph API 错误: " + resp.status;
    }
    const data = await resp.json();
    return data.link.webUrl; // 这里就是 1drv.ms 短链
}

var redata;

window.onload = function() {
    mask_list = document.querySelector(".mask-list");
    mask_name = document.querySelector(".mask-name");
    mask_content = document.querySelector(".mask-content");
    outinfo = document.querySelector(".outinfo");
    outcontent = document.querySelector(".outcontent");

    if (location.protocol !="https:" && location.hostname !="localhost" && location.hostname != "") {
        var goto = confirm("检测到你正在使用http模式，本应用要求使用https模式。\\n是否自动跳转？");
        if (goto) { location.protocol = "https:"; }
    }
    load_mask_local();
}

// OneDrive Picker
function launchOneDrivePicker(action = "query"){
    outinfo.innerHTML = "正在等待API返回数据";
    var odOptions = {
        clientId: "8b57f0ee-14a4-4888-bc40-b10e12a95dd9", // 替换成你的应用ID
        action: action,
        multiSelect: true,
        openInNewWindow: true,
        advanced: {
            redirectUri: "https://shonestone.github.io/GetOneDriveDirectLink/index.html",
            queryParameters: "select=id,name,webUrl,@microsoft.graph.downloadUrl"
        },
        success: function(files) {do_success(files);},
        cancel: function() {do_cancel();},
        error: function(e) {do_error(e);}
    };
    OneDrive.open(odOptions);
}
