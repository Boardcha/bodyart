$('#HomeSlider').slick({
	autoplay: true,
	autoplaySpeed: 5000,
  prevArrow: '<div class="slider-arrow-prev"><i class="fa fa-chevron-left-mdc fa-2x text-white pointer"></i></div>',
  nextArrow: '<div class="slider-arrow-next"><i class="fa fa-chevron-right-mdc fa-2x text-white pointer"></i></div>',
  dots: true,
  customPaging : function(slider, i) {
		var thumb = $(slider.$slides[i]).data('thumb');
		return '<i class="fa fa-primitive-dot fa-lg pointer"></i>';
  }
  /*
	asNavFor: '.home-secondary-slider',
	customPaging : function(slider, i) {
		var thumb = $(slider.$slides[i]).data('thumb');
		return '<i class="fa fa-primitive-dot fa-lg"></i>';
  }, */
  /*
  appendDots: $('.home-slider-dots'),
  responsive: [
    {
      breakpoint: 480,
      settings: {
        dots: false
      }
    }] */
});

$('#NewSlider').slick({
	slidesToShow: 3,
  slidesToScroll: 3,
  prevArrow: '<div class="slider-arrow-prev" style="height:60%"><i class="fa fa-chevron-circle-left fa-2x text-dark pointer"></i></div>',
  nextArrow: '<div class="slider-arrow-next" style="height:60%"><i class="fa fa-chevron-circle-right fa-2x text-dark pointer"></i></div>',
  responsive: [

    {
      breakpoint: 4000,
      settings: {
        slidesToShow: 10,
        slidesToScroll: 10
      }
    },
    {
      breakpoint: 1920,
      settings: {
        slidesToShow: 8,
        slidesToScroll: 8
      }
    },
    {
      breakpoint: 1600,
      settings: {
        slidesToShow: 7,
        slidesToScroll: 7
      }
    },
    {
      breakpoint: 1024,
      settings: {
        slidesToShow: 5,
        slidesToScroll: 5
      }
    },
    {
      breakpoint: 600,
      settings: {
        slidesToShow: 4,
        slidesToScroll: 4
      }
    },
    {
      breakpoint: 480,
      settings: {
        slidesToShow: 3,
        slidesToScroll: 3
      }
    }
    // You can unslick at a given breakpoint now by adding:
    // settings: "unslick"
    // instead of a settings object
  ]
});

$('#testimonials').slick({
	slidesToShow: 3,
  slidesToScroll: 3,
  prevArrow: '<div class="slider-arrow-prev" style="height:60%"><i class="fa fa-chevron-circle-left fa-2x text-dark pointer"></i></div>',
  nextArrow: '<div class="slider-arrow-next" style="height:60%"><i class="fa fa-chevron-circle-right fa-2x text-dark pointer"></i></div>',
  responsive: [

    {
      breakpoint: 4000,
      settings: {
        slidesToShow: 3,
        slidesToScroll: 3
      }
    },
    {
      breakpoint: 1920,
      settings: {
        slidesToShow: 3,
        slidesToScroll: 3
      }
    },
    {
      breakpoint: 1600,
      settings: {
        slidesToShow: 3,
        slidesToScroll: 3
      }
    },
    {
      breakpoint: 1024,
      settings: {
        slidesToShow: 2,
        slidesToScroll: 2
      }
    },
    {
      breakpoint: 600,
      settings: {
        slidesToShow: 1,
        slidesToScroll: 1
      }
    },
    {
      breakpoint: 480,
      settings: {
        slidesToShow: 1,
        slidesToScroll: 1
      }
    }
    // You can unslick at a given breakpoint now by adding:
    // settings: "unslick"
    // instead of a settings object
  ]
});


// Homepage newsletter signup
$("#homepage-newsletter-signup").on("click", function () {
  $("#homepage-newsletter-signup").html('<i class="fa fa-spinner fa-2x fa-spin"></i>');
  $('#homepage_newsletter_email').hide();

  $.ajax({
      method: "post",
      dataType: "json",
      url: "/klaviyo/klaviyo-subscribe-newsletter.asp?email=" + $('#homepage_newsletter_email').val()
      })
      .done(function(json) {
          if ($.isEmptyObject(json)) {
            $("#homepage-newsletter-msg").html('<span class="alert alert-success m-0 p-2">Thanks for signing up!</span>').show();
            $("#homepage-newsletter-signup").hide();
        } 
        if ($.isArray(json)) {
            if ((json[0].id) != "") {
                $("#homepage-newsletter-msg").html('<div class="alert alert-info m-0 p-2">You are already subscribed to our newsletter.</div>').show();
                $("#homepage-newsletter-signup").hide();
            }
        } else {
            if ((json.detail) != "") {
                $("#homepage-newsletter-msg").html('<div class="alert alert-danger m-0 p-2">' + json.detail + '</div>').show().delay(5000).fadeOut("slow");
                $("#homepage-newsletter-signup").html('Sign Up!');
                $('#homepage_newsletter_email').show();
            }

        }     
      })
      .fail(function(json) {			
          $("#homepage-newsletter-msg").html('<div class="alert alert-danger">Website ajax error</div>').show();
          $("#homepage-newsletter-signup").html('Sign Up!');
          $('#homepage_newsletter_email').show();
      })
});
