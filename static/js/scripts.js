$(
    $('.next_event .event').filter(function (event) {
        return Date.parse($(this).data('end')) > Date.now();
    }).sort(function (a,b) {
            return Date.parse($(a).date('start')) - Date.parse($(b).data('start'))
    })[0]
).show();
