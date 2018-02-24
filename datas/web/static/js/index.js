var $ = layui.jquery;
var jQuery = layui.jquery;

Date.prototype.Format = function(fmt) {
    var o = {
        "M+": this.getMonth() + 1, //月份 
        "d+": this.getDate(), //日 
        "h+": this.getHours(), //小时 
        "m+": this.getMinutes(), //分 
        "s+": this.getSeconds(), //秒 
        "q+": Math.floor((this.getMonth() + 3) / 3), //季度 
        "S": this.getMilliseconds() //毫秒 
    };
    if (/(y+)/.test(fmt)) fmt = fmt.replace(RegExp.$1, (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o)
        if (new RegExp("(" + k + ")").test(fmt)) fmt = fmt.replace(RegExp.$1, (RegExp.$1.length == 1) ? (o[k]) : (("00" + o[k]).substr(("" + o[k]).length)));
    return fmt;
}

//右下角工具块
layui.util.fixbar({
    bgcolor: "#393D49"
});

//彩信日志查看
$("#mmsLog").click(function() {
    layer.open({
        type: 2,
        title: "彩信发送日志查看",
        shadeClose: true,
        shade: false,
        maxmin: true, //开启最大化最小化按钮
        skin: "layui-layer-molv",
        area: ["800px", "600px"],
        content: "/logs/mms"
    });
});

//日志查看
$("#logView").click(function() {
    layer.open({
        type: 2,
        title: "调度器日志查看",
        shadeClose: true,
        shade: false,
        maxmin: true, //开启最大化最小化按钮
        skin: "layui-layer-molv",
        area: ["800px", "600px"],
        content: "/logs/AutoReportMms" + new Date().Format("yyyy-MM-dd") + ".log"
    });
});

//重启调度器
$("#taskRestart").click(function() {
    var that = this;
    $(this).attr({
        "disabled": "disabled"
    });
    var loading = layer.load(0, {
        shade: false
    }); //0代表加载的风格，支持0-2
    $.ajax({
        type: "get",
        url: "/api/restart",
        async: false,
        cache: false,
        dataType: "json",
        success: function(data, textStatus) {
            layer.close(loading);
            $(that).removeAttr("disabled");
            layer.msg(data.msg, {
                time: 3 * 1000,
                shift: 0
            });
        },
        error: function(xhr, textStatus, errorThrown) {
            layer.close(loading);
            $(that).removeAttr("disabled");
            layer.msg("重启调度器失败", {
                time: 3 * 1000,
                shift: 0
            });
        }
    });
});

//修改
function taskModify(jobid) {
    layer.open({
        type: 2,
        title: "任务: " + jobid + "修改",
        shadeClose: true,
        shade: false,
        maxmin: true, //开启最大化最小化按钮
        area: ["400px", "300px"],
        content: "/api/task/modify"
    });
}

//查看,暂停,恢复,删除,重启任务-1,2,3,4,5
function taskControl(obj, jobid, type) {
    var loading = layer.load(1, {
        shade: false
    }); //0代表加载的风格，支持0-2
    $.ajax({
        type: "post",
        url: "/api/task/control",
        async: true,
        cache: false,
        dataType: "json",
        data: {
            jobid: jobid,
            type: type
        },
        success: function(data, textStatus) {
            layer.close(loading);
            if (type === 1) {
                //查看
                layer.open({
                    type: 2,
                    title: "任务: " + jobid + "日志查看",
                    shadeClose: true,
                    shade: false,
                    maxmin: true, //开启最大化最小化按钮
                    skin: "layui-layer-molv",
                    area: ["400px", "300px"],
                    content: data.msg
                });
                //                 layer.alert(data.msg, {
                //                     title: "任务: " + jobid,
                //                     skin: "layui-layer-molv",
                //                     closeBtn: 0
                //                 });
            } else {
                if (type === 4) {
                    //删除
                    $(obj).parent().parent().parent().remove();
                }
                layer.msg(data.msg, {
                    time: 3 * 1000,
                    shift: 0
                });
            }
        },
        error: function(xhr, textStatus, errorThrown) {
            layer.close(loading);
            layer.msg("操作失败", {
                time: 3 * 1000,
                shift: 0
            });
        }
    });
}

//获取任务列表
function getTasks(render) {
    var loading = layer.load(0, {
        shade: false
    }); //0代表加载的风格，支持0-2
    $.ajax({
        type: "post",
        url: "/api/tasks",
        async: true,
        cache: false,
        dataType: "json",
        success: function(data, textStatus) {
            var html = render(data); //根据模版渲染数据
            $("#bodyTasks").html(html);
            layer.close(loading);
        },
        error: function(xhr, textStatus, errorThrown) {
            layer.close(loading);
            layer.msg("获取任务列表数据失败", {
                time: 3 * 1000,
                shift: 0
            });
        }
    });
}
//相册层
$.getJSON("/api/qrcode", function(json) {
    if (json.status !== 1) {
        return;
    }
    layer.photos({
        photos: json,
        anim: 5 //0-6的选择，指定弹出图片动画类型，默认随机
    });
});

//获取模版内容
var loading1 = layer.load(0, {
    shade: false
}); //0代表加载的风格，支持0-2
var render = template.compile($("#artTasks").html()); //编译模版文件
getTasks(render);
setInterval(function() {
    getTasks(render);
}, 15000); //15秒钟
layer.close(loading1);