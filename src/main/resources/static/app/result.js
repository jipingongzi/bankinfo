window.alert = function(msg){
    var config = {
        fontSize:'12px',
        backgroundColor:'#fa8072',
        fontColor:'#FFFFFF',
        borderRadius:'20px',
        border:'1px #696969',
        padding:'15px 15px'
    };
    var $alertPanel = $("#alertPanel263");
    if($alertPanel.length>0){
        $alertPanel.find("span").html(msg);
    }else{
        var html = new Array();
        html.push('<div id="alertPanel263" style="display:none;position:fixed;width:100%;bottom:35%;text-align:center;z-index:9999999">')
        html.push('<span style="padding:'+config.padding+';border:'+config.border+';font-size:'+config.fontSize+';font-weight: bold;background-color:'+config.backgroundColor+';color:'+config.fontColor+';border-radius:'+config.borderRadius+';-moz-border-radius:'+config.borderRadius+';-webkit-border-radius:'+config.borderRadius+'">')
        html.push(msg)
        html.push('</span></div>');
        $alertPanel = $(html.join(''));
        $("body").append($alertPanel);
    }
    $alertPanel.stop().fadeIn('fast').delay(2000).fadeOut('slow');
};