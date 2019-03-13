$('.wminimize').click(function (e) {
    e.preventDefault();
    var $wcontent = $(this).parent().parent().next('.widget-content');
    if ($wcontent.is(':visible')) {
        $(this).children('i').removeClass('icon-circle-arrow-up');
        $(this).children('i').addClass('icon-circle-arrow-down');
    }
    else {
        $(this).children('i').removeClass('icon-circle-arrow-down');
        $(this).children('i').addClass('icon-circle-arrow-up');
    }
    $wcontent.toggle(500);
});
$('.wclose').click(function (e) {
    e.preventDefault();
    var $wbox = $(this).parent().parent().parent();
    $wbox.hide(100);
});