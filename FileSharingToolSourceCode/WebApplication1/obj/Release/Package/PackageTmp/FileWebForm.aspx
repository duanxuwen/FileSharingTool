<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="FileWebForm.aspx.cs" Inherits="WebApplication1.FileWebForm" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <link href="FontAwesome/css/font-awesome.css" rel="stylesheet" />
    <link href="FontAwesome/css/font-awesome.min.css" rel="stylesheet" />

    <link href="style/bootstrap.css" rel="stylesheet" />
    <link href="style/style.css" rel="stylesheet" />

    <style type="text/css">
        .jqstooltip {
            position: absolute;
            left: 0px;
            top: 0px;
            visibility: hidden;
            background: rgb(0, 0, 0) transparent;
            background-color: rgba(0,0,0,0.6);
            filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#99000000, endColorstr=#99000000);
            -ms-filter: "progid:DXImageTransform.Microsoft.gradient(startColorstr=#99000000, endColorstr=#99000000)";
            color: white;
            font: 10px arial, san serif;
            text-align: left;
            white-space: nowrap;
            padding: 5px;
            border: 1px solid white;
            z-index: 10000;
        }

        .jqsfield {
            color: white;
            font: 10px arial, san serif;
            text-align: left;
        }

        .widget-head i {
            position: relative;
            width: 60px;
            height: 60px;
            border-radius: 50%;
            border: 1px solid #2b92d4;
            color: #fff;
            text-align: center;
            overflow: hidden;
            background-image: -webkit-gradient(linear, left top, left bottom, from(#6cc3fe), to(#21a1d0));
            -webkit-animation-timing-function: ease-in-out;
            -webkit-animation-name: breathe;
            -webkit-animation-duration: 2700ms;
            -webkit-animation-iteration-count: infinite;
            -webkit-animation-direction: alternate;
        }

        @-webkit-keyframes breathe {
            0% {
                opacity: .2;
                box-shadow: 0 1px 2px rgba(255,255,255,0.1);
            }

            100% {
                opacity: 1;
                border: 1px solid rgba(59,235,235,1);
                box-shadow: 0 1px 30px rgba(59,255,255,1);
            }
        }
    </style>
    <script>

        function Change(src) {
            //$("#FilePath").html(src);
            //document.getElementById('FilePath').innerText = src;
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <asp:FileUpload ID="File" runat="server" Style="display: none" />

        <asp:Repeater ID="Repeater" runat="server" OnItemCommand="Repeater_ItemCommand">
            <ItemTemplate>

                <div class="widget wblack">
                    <div class="widget-head">
                        <div class="pull-left"><%#Eval("Title")%></div>
                        <div class="widget-icons pull-right" style="width: 50px">
                            <a href="#" class="wminimize"><i class="icon-circle-arrow-down"></i></a>
                            <a href="#" class="wclose"><i class="icon-remove-sign"></i></a>
                        </div>
                        <div class="clearfix"></div>
                    </div>
                    <div class="widget-content" style="display: none;">
                        <iframe id="IFRAME1" runat="server" height="600" style="width: 100%" src='<%#Eval("Address")%>'></iframe>
                        <div class="widget-foot">
                            <a href="#" class="btn btn-danger" onclick="document.getElementById('File').click();">选择文件</a>
                            <asp:Button Text="上传" class="btn btn-danger" CommandName="上传" runat="server" />
                            <asp:Button Text="下载" class="btn btn-warning"  CommandArgument='<%#Eval("Address")%>' CommandName="下载" runat="server" />
                            <asp:Button Text="删除" class="btn btn-warning"  CommandArgument='<%#Eval("Address")%>' CommandName="删除" runat="server" />
                            <asp:Button Text="合并" class="btn btn-warning"  CommandArgument='<%#Eval("Address")%>' CommandName="合并" runat="server" />
                        </div>
                    </div>
                </div>

            </ItemTemplate>
        </asp:Repeater>
    </form>
    <script src="js/jquery_005.js"></script>
    <script src="js/js.js"></script>
</body>
</html>
