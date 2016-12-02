var old_keyword;

$('.login-navigation li').click(function() {
    $('.login-error-message').hide();
});

$('.query_form').submit(function(event) {
    event.preventDefault();
    var keyword = $('#q').val();
    if (keyword != old_keyword) {
        old_keyword = keyword;
        $.ajax({
            method: 'POST',
            url: '/retrieve_post',
            data: {
                'q': keyword,
                "csrfmiddlewaretoken": csrf_token,
            },
            success: function(ret) {
                if (ret['status'] == 'complete') {
                    window.location.href = '/post/' + keyword;
                } else {
                    $('.query-preloader').removeClass('hidden');
                    $('#btn-query').prop('disabled', true);

                    setTimeout(function() {
                        window.location.href = '/post/' + keyword + '/_';
                    }, 5000);
                }
            }
        });
    }
})

function show_comments(post_id) {
    $.ajax({
        method: 'POST',
        url: '/retrieve_comment',
        data: {
            'post_id': post_id,
            "csrfmiddlewaretoken": csrf_token,
        },
        success: function(ret) {
            if (ret['status'] == 'complete') {
                window.location.href = '/comment/' + post_id;
            } else {
                $('.page-loader').show();

                setTimeout(function() {
                    window.location.href = '/comment/' + post_id + '/_';
                }, 5000);
            }
        }
    });
};
