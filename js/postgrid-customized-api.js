let addressDivs;
let PID = 0;
/**
 * @type {{[key:number|string]:(addr:any)=>any}}
 */
let customAddressCompletedCallbacks = {};

// ClickFunnel Address Form
{
    let shippingDivs = document.querySelectorAll(
        '.elOS1Shipping, .elShippingForm'
    );

    for (const div of shippingDivs) {
        div.setAttribute('data-pg-verify', '');
    }

    addressDivs = document.getElementsByName('shipping_address');
    const cityDivs = document.getElementsByName('shipping_city');
    const stateDivs = document.getElementsByName('shipping_state');
    const zipDivs = document.getElementsByName('shipping_zip');
    const countryDivs = document.getElementsByName('shipping_country');

    for (const div of addressDivs) {
        div.setAttribute('data-pg-full-address', '');
    }
    
	for (const div of addressDivs) {
        div.setAttribute('data-pg-address-line1', '');
    }

    for (const div of cityDivs) {
        div.setAttribute('data-pg-city', '');
    }

    for (const div of stateDivs) {
        div.setAttribute('data-pg-prov', '');
    }
	
    for (const div of stateDivs) {
        div.setAttribute('data-pg-prov2', '');
    }

    for (const div of stateDivs) {
        div.setAttribute('data-pg-prov3', '');
    }	

    for (const div of zipDivs) {
        div.setAttribute('data-pg-pc', '');
    }

    for (const div of countryDivs) {
        div.setAttribute('data-pg-country', '');
    }
}

let usingGravityForm = false;

// GravityForm Address Form
{
    let shippingDivs = document.querySelectorAll('.ginput_container_address');

    for (const div of shippingDivs) {
        div.setAttribute('data-pg-verify', '');
    }

    addressDivs = document.querySelectorAll('.address_line_1 > input');
    const address2Divs = document.querySelectorAll('.address_line_2 > input');
    const cityDivs = document.querySelectorAll('.address_city > input');
    const stateDivs = document.querySelectorAll('.address_state > input');
    const zipDivs = document.querySelectorAll('.address_zip > input');
    const countryDivs = document.querySelectorAll(
        '.address_country > input, .address_country > select'
    );

    for (const div of addressDivs) {
        div.setAttribute('data-pg-full-address', '');
    }
	
    for (const div of addressDivs) {
        div.setAttribute('data-pg-address-line1', '');
    }

    for (const div of address2Divs) {
        div.setAttribute('data-pg-address-line2', '');
    }

    for (const div of cityDivs) {
        div.setAttribute('data-pg-city', '');
    }

    for (const div of stateDivs) {
        div.setAttribute('data-pg-prov', '');
    }

    for (const div of stateDivs) {
        div.setAttribute('data-pg-prov2', '');
    }

    for (const div of stateDivs) {
        div.setAttribute('data-pg-prov3', '');
    }
	
    for (const div of zipDivs) {
        div.setAttribute('data-pg-pc', '');
    }

    for (const div of countryDivs) {
        div.setAttribute('data-pg-country', '');
    }

    if (shippingDivs.length) {
        usingGravityForm = true;
    }
}

var baseUrl =
    (document.querySelector('[data-pg-base-url]') &&
        document
            .querySelector('[data-pg-base-url]')
            .getAttribute('data-pg-base-url')) ||
    'https://api.postgrid.com';

// ***********************************************************************************************************

// Retrieved from https://gist.github.com/doxxx/8987233
if (!Element.prototype.scrollIntoViewIfNeeded) {
    Element.prototype.scrollIntoViewIfNeeded = function (centerIfNeeded) {
        centerIfNeeded = arguments.length === 0 ? true : !!centerIfNeeded;

        var parent = this.parentNode,
            parentComputedStyle = window.getComputedStyle(parent, null),
            parentBorderTopWidth = parseInt(
                parentComputedStyle.getPropertyValue('border-top-width')
            ),
            parentBorderLeftWidth = parseInt(
                parentComputedStyle.getPropertyValue('border-left-width')
            ),
            overTop = this.offsetTop - parent.offsetTop < parent.scrollTop,
            overBottom =
                this.offsetTop -
                    parent.offsetTop +
                    this.clientHeight -
                    parentBorderTopWidth >
                parent.scrollTop + parent.clientHeight,
            overLeft = this.offsetLeft - parent.offsetLeft < parent.scrollLeft,
            overRight =
                this.offsetLeft -
                    parent.offsetLeft +
                    this.clientWidth -
                    parentBorderLeftWidth >
                parent.scrollLeft + parent.clientWidth;

        if (centerIfNeeded) {
            if (overTop || overBottom) {
                parent.scrollTop =
                    this.offsetTop -
                    parent.offsetTop -
                    parent.clientHeight / 2 -
                    parentBorderTopWidth +
                    this.clientHeight / 2;
            }

            if (overLeft || overRight) {
                parent.scrollLeft =
                    this.offsetLeft -
                    parent.offsetLeft -
                    parent.clientWidth / 2 -
                    parentBorderLeftWidth +
                    this.clientWidth / 2;
            }
        } else {
            if (overTop) {
                parent.scrollTop =
                    this.offsetTop - parent.offsetTop - parentBorderTopWidth;
            }

            if (overBottom) {
                parent.scrollTop =
                    this.offsetTop -
                    parent.offsetTop -
                    parentBorderTopWidth -
                    parent.clientHeight +
                    this.clientHeight;
            }

            if (overLeft) {
                parent.scrollLeft =
                    this.offsetLeft - parent.offsetLeft - parentBorderLeftWidth;
            }

            if (overRight) {
                parent.scrollLeft =
                    this.offsetLeft -
                    parent.offsetLeft -
                    parentBorderLeftWidth -
                    parent.clientWidth +
                    this.clientWidth;
            }
        }
    };
}

// If line 1 address is div, append a textbox and if it is a textbox then wrap a
// div with position relative

var wrap = function (toWrap, wrapper = undefined) {
    wrapper = wrapper || document.createElement('div');
    // wrapper.setAttribute('data-pg-address','')
    wrapper.setAttribute('class', 'pg-input-wrapper');
    toWrap.parentNode.insertBefore(wrapper, toWrap);
    return wrapper.appendChild(toWrap);
};

for (var elem of document.querySelectorAll(
    '[data-pg-verify] [data-pg-address-line1]'
)) {
    var inputBox;
    if (
        elem.tagName.toLocaleLowerCase() === 'input' ||
        elem.tagName.toLocaleLowerCase() === 'textarea'
    ) {
        inputBox = elem;
        wrap(elem);
    } else {
        inputBox = document.createElement('input');
        inputBox.setAttribute('type', 'text');
        var attributes = elem.attributes;
        // Move all attributes from parent element (div) to input box
        for (var i = attributes.length - 1; i >= 0; i--) {
            var attribute = attributes[i];
            inputBox.setAttribute(attribute.name, attribute.value);
            elem.removeAttribute(attribute.name);
        }
        elem.setAttribute('class', 'pg-input-wrapper');
        elem.appendChild(inputBox);
    }
    inputBox.setAttribute('type', 'text');
    inputBox.setAttribute('autocomplete', 'off');
}

let forms = document.querySelectorAll('[data-pg-verify]');

var currentFocus;
var currentTimeout;

var debounceTime = document.querySelector('[data-pg-debounce-time]')
    ? document
          .querySelector('[data-pg-debounce-time]')
          .getAttribute('data-pg-debounce-time')
    : 100;

if (isNaN(debounceTime) || debounceTime === '') {
    debounceTime = 100;
} else {
    debounceTime = Number.parseInt(debounceTime);
}

for (let i = 0; i < forms.length; i++) {
    let config = {
        elements: {
            form: forms[i],
            countrySelected: forms[i].querySelector('[data-pg-select-country]'),
            line1: forms[i].querySelectorAll('[data-pg-address-line1]'),
			fullAddress: forms[i].querySelectorAll('[data-pg-full-address]'),
            line1Orig: forms[i].querySelectorAll('[data-pg-orig-address-line1]'),
            line1Msg: forms[i].querySelectorAll('[data-pg-address-line1-message]'),
            line2: forms[i].querySelector('[data-pg-address-line2]'),
            line2Msg: forms[i].querySelector('[data-pg-address-line2-message]'),
            city: forms[i].querySelector('[data-pg-city]'),
            cityMsg: forms[i].querySelector('[data-pg-city-message]'),
            prov: forms[i].querySelector('[data-pg-prov]'),
            provMsg: forms[i].querySelector('[data-pg-prov-message]'),
            prov2: forms[i].querySelector('[data-pg-prov2]'),
            provMsg: forms[i].querySelector('[data-pg-prov2-message]'),
            prov3: forms[i].querySelector('[data-pg-prov3]'),
            provMsg: forms[i].querySelector('[data-pg-prov3-message]'),			
            pc: forms[i].querySelector('[data-pg-pc]'),
            pcMessage: forms[i].querySelector('[data-pg-pc-message]'),
            country: forms[i].querySelector('[data-pg-country]'),
            countryMessage: forms[i].querySelector('[data-pg-country-message]'),
            status: document.querySelector('[data-pg-status]'),
            errorBox: document.querySelector('[data-pg-generic-message]'),
            invalidBox: [],
        },
        apis: {
            verify: baseUrl + '/v1/addver/verifications',
            autocomplete: baseUrl + '/v1/addver/completions',
            intlVerify: baseUrl + '/v1/intl_addver/verifications',
            intlAutocomplete: baseUrl + '/v1/intl_addver/completions',
        },
        apiKey: document.querySelector('[data-pg-key]')
            ? document
                  .querySelector('[data-pg-key]')
                  .getAttribute('data-pg-key')
            : null,
        isInternational: document.querySelector('[data-pg-international]')
            ? document
                  .querySelector('[data-pg-international]')
                  .getAttribute('data-pg-international') === 'true'
            : false,
        countryFilter: document.querySelector('[data-pg-country-filter]')
            ? document
                  .querySelector('[data-pg-country-filter]')
                  .getAttribute('data-pg-country-filter')
            : null,
        skipVerification: document.querySelector('[data-pg-skip-verification]')
            ? document
                  .querySelector('[data-pg-skip-verification]')
                  .getAttribute('data-pg-skip-verification') === 'true'
            : true,
        fullAutocomplete: document.querySelector('[data-pg-full-autocomplete]')
            ? document
                  .querySelector('[data-pg-full-autocomplete]')
                  .getAttribute('data-pg-full-autocomplete') === 'true'
            : false,
        fullAddressLine1: document.querySelector('[data-pg-full-address-line1]')
            ? document
                  .querySelector('[data-pg-full-address-line1]')
                  .getAttribute('data-pg-full-address-line1') === 'true'
            : true,
        useProvinceCodes: document.querySelector('[data-pg-use-province-codes]')
            ? document
                  .querySelector('[data-pg-use-province-codes]')
                  .getAttribute('data-pg-use-province-codes') === 'true'
            : false,
    };
    // ***************************************** Autocomplete line 1 **************************************************

    if (config.elements.fullAddress && config.elements.fullAddress.length > 0) {
        if (!config.skipVerification) {
            var invalidBox = document.createElement('input');
            invalidBox.setAttribute('required', 'true');
            invalidBox.setAttribute('id', 'pg-invalid-box');
            invalidBox.setAttribute('hidden', 'true');
            config.elements.invalidBox.push(invalidBox);
            // When user tries to submit
            invalidBox.addEventListener('invalid', function (ev) {
                verifyFields(ev, config);
            });

            config.elements.fullAddress[0].parentNode.insertBefore(
                invalidBox,
                config.elements.fullAddress[0]
            );
        }

        for (var fullAddress of config.elements.fullAddress) {
            autocomplete(fullAddress, config);
            fullAddress.addEventListener('input', function () {
                onInput(config);
            });
        }
    }
    if (config.elements.line2) {
        config.elements.line2.addEventListener('input', function () {
            onInput(config);
        });
    }
    if (config.elements.city) {
        config.elements.city.addEventListener('input', function () {
            onInput(config);
        });
    }
    if (config.elements.prov) {
        config.elements.prov.addEventListener('input', function () {
            onInput(config);
        });
    }
    if (config.elements.prov2) {
        config.elements.prov2.addEventListener('input', function () {
            onInput(config);
        });
    }
    if (config.elements.prov3) {
        config.elements.prov3.addEventListener('input', function () {
            onInput(config);
        });
    }	
    if (config.elements.pc) {
        config.elements.pc.addEventListener('input', function () {
            onInput(config);
        });
    }
    if (config.elements.country) {
        config.elements.country.addEventListener('input', function () {
            onInput(config);
        });
    }

    const standardCountriesAbbreviations = new Set(['us', 'ca']);

    if (
        config.countryFilter &&
        !standardCountriesAbbreviations.has(config.countryFilter.toLowerCase())
    ) {
        config.countryFilter = null;
    }
}

function autocomplete(inp, config) {
    inp.addEventListener('input', function (ev) {
        // Search text in line 1
        PID += 1;
        currentTimeout = efficientSearch(ev, config, PID);
    });

    inp.addEventListener('keydown', function (ev) {
        chooseOption(ev);
    });

    /*execute a function when someone clicks in the document:*/
    document.addEventListener('click', function (e) {
        if (e.target === inp) {
            return;
        }
        closeAllLists();
    });
}

// ****************************************** Search List
// *********************************************************

function debounce(func, wait, immediate) {
    var timeout;
    return function () {
        var context = this,
            args = arguments;

        var later = function later() {
            timeout = null;
            if (!immediate) func.apply(context, args);
        };

        var callNow = immediate && !timeout;
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
        if (callNow) func.apply(context, args);

        return timeout;
    };
}

const docBody = document.body;
const docEle = document.documentElement;

const pageHeight = Math.max(
    docBody.scrollHeight,
    docBody.offsetHeight,
    docEle.clientHeight,
    docEle.scrollHeight,
    docEle.offsetHeight
);

const menuHeight = pageHeight / 2.5;
let inputHeight = 0;
let position = 0;
let optionHeight = 0;

var efficientSearch = debounce(function (ev, config, pid) {
    var text = ev.target.value;
    if (!text) {
        return;
    }

    function setValueAndErrors(addr, section) {

        const parsedAddress = config.isInternational ? addr : addr.address;
        const errors = addr.errors;
        const value = config.isInternational
            ? parsedAddress.formattedAddress.replaceAll('\n', ' ')
            : parseAutoCompleteAddressToString(parsedAddress);
        ev.target.value = value;
		
			console.log("addr");
	console.log(parsedAddress);
	
        // Show response in div
        for (let i = 0; i < config.elements.line1.length; i++) {
            if (config.fullAddressLine1) {
                const elem = config.elements.line1[i];

                setValueIfExists(
                    parsedAddress,
                    config.isInternational ? 'line1' : 'address',
                    config.elements.line1[i]
                );				

                if (config.elements.line1Orig.length > i) {
                    setValueIfExists(
                        parsedAddress,
                        config.isInternational ? 'line1' : 'address',
                        config.elements.line1Orig[i]
                    );
                }
            } else {
                setValueIfExists(
                    parsedAddress,
                    config.isInternational ? 'line1' : 'address',
                    config.elements.line1[i]
                );
            }
        }
        if (config.isInternational) {
            setValueIfExists(parsedAddress, 'line2', config.elements.line2);
        }
        setValueIfExists(parsedAddress, 'city', config.elements.city);

        setValueIfExists(
            parsedAddress,
            config.isInternational ? 'postalOrZip' : 'pc',
            config.elements.pc
        );

        if (config.isInternational) {
            if (usingGravityForm) {
                // Complete country name for gravity form
                setValueIfExists(
                    parsedAddress,
                    'country',
                    config.elements.country
                );
            } else {
                setValueIfExists(
                    parsedAddress,
                    'countryCode',
                    config.elements.country
                );
            }
        } else {
            if (usingGravityForm) {
                if (parsedAddress.country.toLowerCase() === 'ca') {
                    parsedAddress.country = 'Canada';
                } else {
                    parsedAddress.country = 'United States';
                }
            } else {
                parsedAddress.country = parsedAddress.country.toUpperCase();
            }		
            setValueIfExists(parsedAddress, 'country', config.elements.country);
			
			setValueIfExists(
				parsedAddress,
				config.isInternational
					? config.useProvinceCodes
						? 'provinceCode'
						: 'provinceOrState'
					: 'prov',
				config.elements.prov
			);

			setValueIfExists(
				parsedAddress,
				config.isInternational
					? config.useProvinceCodes
						? 'provinceCode'
						: 'provinceOrState'
					: 'prov',
				config.elements.prov2
			);

			setValueIfExists(
				parsedAddress,
				config.isInternational
					? config.useProvinceCodes
						? 'provinceCode'
						: 'provinceOrState'
					: 'prov',
				config.elements.prov3
			);			
        }
		
		setSelectedAddress(parsedAddress, section);

        errors && setErrorIfExists(errors, config);

        Object.keys(customAddressCompletedCallbacks).forEach((key) => {
            const f = customAddressCompletedCallbacks[key];
            f(addr);
        });
    }

    const countrySelected = config.isInternational
        ? config.elements.country
            ? getElementValue(config.elements.country)
            : ''
        : null;

    const execGetList = (
        txt,
        country,
        config,
        isPostRequest,
        advanced,
        container
    ) => {
        position = 0;
        optionHeight = 0;
        getList(txt, country, config, isPostRequest, advanced, container)
            .then(function (list) {
                if (pid !== PID) {
                    return;
                }
                var b,
                    val = list;
                closeAllLists();
                if (!val) {
                    return false;
                }
                currentFocus = -1; // Set focused element in list to none
                a = document.createElement('div');
                a.setAttribute('class', 'pg-autocomplete-list');
                a.setAttribute(
                    'style',
                    'width: ' +
                        ev.target.clientWidth +
                        'px; ' +
                        'max-height: ' +
                        menuHeight +
                        'px; ' +
                        'overflow-y: scroll;'
                );
                ev.target.parentNode.appendChild(a);
                inputHeight = ev.target.offsetHeight;
                for (let i = 0; i < list.length; i++) {
                    const option = config.isInternational
                        ? list[i]
                        : config.fullAutocomplete
                        ? list[i].address
                        : list[i].preview;
                    let val = config.isInternational
                        ? parseAutoCompleteIntlAddressToHighlightedString(
                              option
                          )
                        : parseAutoCompleteAddressToString(option);
                    b = document.createElement('div');
                    b.setAttribute('data-val', val);
                    b.innerHTML += val;
                    b.innerHTML += "<input type='hidden' value='" + val + "'>";
                    const ind = config.isInternational ? option.id : i;

                    b.addEventListener('click', function (e, index = ind) {
						var section = "";
						console.log($(this).parent().parent().parent().attr("id"));
						if($(this).parent().parent().parent().attr("id") == "billing-address-autocomplete")
							section = "billing";
						if($(this).parent().parent().parent().attr("id") == "shipping-address-autocomplete")
							section = "shipping";							
                        if (config.fullAutocomplete) {
                            setValueAndErrors(list[i], section);
                        } else {
                            if (config.isInternational) {
                                if (list[i].type === 'Address') {
                                    getIntlAutocomplete(index, config).then(
                                        function (addr) {
                                            setValueAndErrors(addr, section);
                                        }
                                    );
                                } else {
                                    execGetList(
                                        null,
                                        null,
                                        config,
                                        config.fullAutocomplete,
                                        true,
                                        list[i].id
                                    );
                                }
                            } else {
                                getAutocomplete(text, index, config).then(
                                    function (addr) {
                                        setValueAndErrors(addr, section);
                                    }
                                );
                            }
                        }
                        closeAllLists();
                    });
                    a.appendChild(b);
                }
            })
            .catch(function (err) {
                return console.warn(err);
            });
    };

    execGetList(text, countrySelected, config, config.fullAutocomplete);
}, debounceTime); // debounce time

function setSelectedAddress(option, section) {
	
	var address = (option.address ? option.address + '<br/>' : '') +
		(option.city ? option.city + '<br/>' : '') +
		(option.prov ? option.prov + '<br/>' : '') +
		(option.pc ? option.pc + '<br/>' : '') +
		(option.country ? option.country + '' : '');
	
			
	var content = '<div class="alert alert-secondary alert-dismissible fade show" role="alert">' + 
	'  <div id="selected-' + section + '-address-content" class="m-2"><div class="mb-2 font-weight-bold">' + ((section == 'shipping') ? '<i class="fa fa-shipping-fast fa-lg mr-2"></i> ':'') + '<span style="text-transform: uppercase;">' + section + ' ADDRESS</span></div><div style="line-height:22px;">' + address + '</div></div>' +
	'  <button type="button"  class="close" id="btn-edit-' + section + '-address" style="right:20px;padding: 7px 11px 7px 11px;margin-right:16px">' + 
	'	<img src="/images/edit.svg" style="height:14px;width:14px;vertical-align:initial;" />'  +
	'  </button>' +	
	'  <button id="' + section + '-bubble-close" type="button" class="close" data-dismiss="alert" aria-label="Close" style="padding: 7px 11px 7px 11px" onClick="$(\'#' + section + '-full-address\').val(\'\');$(\'#' + section + '-full-address\').focus();clearAddressInputs(\'' + section + '\')">' + 
	'	<span aria-hidden="true">&times;</span>' +
	'  </button>' +
	'</div>';
	
    $('#selected-' + section + '-address').html(content);
	$('#selected-' + section + '-address').show();	
	$('#' + section + '-country').change();	
	
	if($("input[name='shipping-same-billing']").is(':checked') && section == 'shipping')
		$("input[name='shipping-same-billing']").change();
	
}

function parseAutoCompleteAddressToString(option) {
    return (
        '' +
        (option.address ? option.address + ' ' : '') +
        (option.city ? option.city + ' ' : '') +
        (option.prov ? option.prov + ' ' : '') +
        (option.pc ? option.pc + ' ' : '')
    );
}

function parseAutoCompleteIntlAddressToString(option) {
    return (
        '' +
        (option.text ? option.text + ' ' : '') +
        (option.description ? option.description + ' ' : '')
    );
}

function parseAutoCompleteIntlAddressToHighlightedString(option) {
    let textVal = '';
    let descriptionVal = '';
    let lastIndex = 0;
    let parts = null;

    if (option.highlight) {
        parts = option.highlight.split(';');
    }

    if (parts && parts[0]) {
        const textSegments = parts[0].split(',');

        for (const segment of textSegments) {
            const indexes = segment.split('-');
            textVal +=
                option.text.slice(lastIndex, indexes[0]) +
                '<b>' +
                option.text.slice(indexes[0], indexes[1]) +
                '</b>';
            lastIndex = indexes[1];
        }

        textVal += option.text.slice(lastIndex);
    } else {
        textVal = option.text;
    }

    lastIndex = 0;

    if (parts && parts[1]) {
        const descriptionSegments = parts[1].split(',');

        for (const segment of descriptionSegments) {
            const indexes = segment.split('-');
            descriptionVal +=
                option.description.slice(lastIndex, indexes[0]) +
                '<b>' +
                option.description.slice(indexes[0], indexes[1]) +
                '</b>';
            lastIndex = indexes[1];
        }

        descriptionVal += option.description.slice(lastIndex);
    } else {
        descriptionVal = option.description;
    }

    return (
        (textVal ? textVal + ' ' : '') +
        (descriptionVal ? descriptionVal + ' ' : '')
    );
}

function getAutocomplete(txt, index, config) {
    return new Promise(function (resolve, reject) {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', config.apis.autocomplete + '?index=' + index);

        xhr.setRequestHeader('Content-Type', 'application/json');
        if (config.apiKey) {
            xhr.setRequestHeader('X-API-KEY', config.apiKey);
        }
        xhr.onload = function () {
            var resp = JSON.parse(xhr.responseText);
            if (xhr.status === 200) {
                resolve(resp.data);
            } else {
                reject(resp);
            }
        };

        xhr.send(
            JSON.stringify({
                partialStreet: txt,
                countryFilter: config.countryFilter
                    ? config.countryFilter
                    : undefined,
            })
        );
    });
}

function getIntlAutocomplete(id, config) {
    return new Promise(function (resolve, reject) {
        var xhr = new XMLHttpRequest();
        xhr.open('POST', config.apis.intlAutocomplete);
        xhr.setRequestHeader('Content-Type', 'application/json');
        if (config.apiKey) {
            xhr.setRequestHeader('X-API-KEY', config.apiKey);
        }
        xhr.onload = function () {
            var resp = JSON.parse(xhr.responseText);
            if (xhr.status === 200) {
                resolve(resp.data);
            } else {
                reject(resp);
            }
        };

        xhr.send(
            JSON.stringify({
                id: id,
            })
        );
    });
}

function getList(
    txt,
    country,
    config,
    isPostRequest = false,
    advanced = false,
    container
) {
    return new Promise(function (resolve, reject) {
        var xhr = new XMLHttpRequest();
        if (isPostRequest) {
            xhr.open(
                'POST',
                config.isInternational
                    ? config.apis.intlAutocomplete
                    : config.apis.autocomplete
            );
        } else {
            const query = config.isInternational
                ? advanced
                    ? `?container=${encodeURIComponent(
                          container
                      )}&advanced=${true}`
                    : `?partialStreet=${encodeURIComponent(
                          txt
                      )}&countriesFilter=${encodeURIComponent(country)}`
                : `?partialStreet=${encodeURIComponent(txt)}` +
                  (config.countryFilter
                      ? `&countryFilter=${encodeURIComponent(
                            config.countryFilter
                        )}`
                      : '');
            xhr.open(
                'GET',
                (config.isInternational
                    ? config.apis.intlAutocomplete
                    : config.apis.autocomplete) + query
            );
        }

        xhr.setRequestHeader('Content-Type', 'application/json');
        if (config.apiKey) {
            xhr.setRequestHeader('X-API-KEY', config.apiKey);
        }

        xhr.onload = function () {
            if (xhr.status === 200) {
                var resp = JSON.parse(xhr.responseText);
                var response = resp.data;

                resolve(response);
            } else {
                var resp = JSON.parse(xhr.responseText);
                reject(resp);
            }
        };
        if (isPostRequest) {
            xhr.send(JSON.stringify({ partialStreet: txt }));
        } else {
            xhr.send();
        }
    });
}

// ************************************* Choose option on down or up key arrow
// ************************************

function chooseOption(e) {
    const maxOptions =
        Math.floor(parseFloat(menuHeight) / parseFloat(inputHeight)) - 1;
    let menu = document.querySelector('.pg-autocomplete-list');
    var x = document.querySelectorAll('.pg-autocomplete-list')[0];
    if (!x) {
        return;
    }
    x = x.getElementsByTagName('div');

    if (position < 1) {
        position = 0;
    } else if (position >= maxPosition) {
        position = maxPosition;
    }

    if (e.keyCode == 40) {
        /*If the arrow DOWN key is pressed,
        increase the currentFocus variable:*/
        currentFocus++;

        if (currentFocus >= x.length) {
            currentFocus = x.length - 1;
        }

        x[currentFocus].scrollIntoViewIfNeeded(false);

        addActive(x);
    } else if (e.keyCode == 38) {
        // up
        /*If the arrow UP key is pressed,
        decrease the currentFocus variable:*/
        currentFocus--;

        if (currentFocus < 1) {
            currentFocus = 0;
        }

        x[currentFocus].scrollIntoViewIfNeeded(false);

        addActive(x);
    } else if (e.keyCode == 13) {
        /*If the ENTER key is pressed, prevent the form from being submitted,*/
        e.preventDefault();

        if (currentTimeout) {
            clearTimeout(currentTimeout);
        }

        if (x) {
            if (currentFocus > -1) {
                /*and simulate a click on the "active" item:*/
                x[currentFocus].click();
            } else if (x.length === 1) {
                x[0].click();
            }
        }
    }
}

function addActive(x) {
    /*a function to classify an item as "active":*/
    if (!x) return false;
    /*start by removing the "active" class on all items:*/
    removeActive(x);

    if (currentFocus >= x.length) currentFocus = x.length - 1;
    if (currentFocus < 1) currentFocus = 0;

    /*add class "autocomplete-active":*/
    x[currentFocus].setAttribute('data-pg-address-active', '');
}

function removeActive(x) {
    /*a function to remove the "active" class from all autocomplete items:*/
    for (var i = 0; i < x.length; i++) {
        x[i].removeAttribute('data-pg-address-active');
    }
}

function closeAllLists() {
    var x = document.querySelectorAll('.pg-autocomplete-list');
    for (var i = 0; i < x.length; i++) {
        x[i].parentNode.removeChild(x[i]);
    }
}

// ****************************************** On Form Submit
// ******************************************************

function verify(config) {
    return new Promise(function (resolve, reject) {
        if (!config.elements.line1.length === 0) {
            return;
        }

        var address = {
            line1: getElementValue(config.elements.line1[0]),
        };
        if (config.elements.line2) {
            address.line2 = getElementValue(config.elements.line2);
        }
        if (config.elements.city) {
            address.city = getElementValue(config.elements.city);
        }
        if (config.elements.prov) {
            address.provinceOrState = getElementValue(config.elements.prov);
        }	
        if (config.elements.pc) {
            address.postalOrZip = getElementValue(config.elements.pc);
        }
        if (config.elements.country) {
            address.country = getElementValue(config.elements.country);
        }
        var xhr = new XMLHttpRequest();
        xhr.open(
            'POST',
            config.isInternational ? config.apis.intlVerify : config.apis.verify
        );
        xhr.setRequestHeader('Content-Type', 'application/json');
        if (config.apiKey) {
            xhr.setRequestHeader('X-API-KEY', config.apiKey);
        }
        xhr.onload = function () {
            var resp = JSON.parse(xhr.responseText);
            if (xhr.status === 200) {
                resolve(resp.data);
            } else {
                reject(resp);
            }
        };
        xhr.send(JSON.stringify({ address }));
    });
}

function onInput(config) {
    for (let i = 0; i < config.elements.invalidBox.length; i++) {
        if (config.elements.invalidBox[i]) {
            config.elements.invalidBox[i].setAttribute('required', 'true');
        }
    }
    for (let i = 0; i < config.elements.line1Msg.length; i++) {
        config.elements.line1Msg[i].innerHTML = '';
    }
    if (config.elements.errorBox) {
        config.elements.errorBox.innerHTML = '';
    }
    if (config.elements.cityMsg) {
        config.elements.cityMsg.innerHTML = '';
    }
    if (config.elements.provMsg) {
        config.elements.provMsg.innerHTML = '';
    }
    if (config.elements.prov2Msg) {
        config.elements.prov2Msg.innerHTML = '';
    }
    if (config.elements.prov3Msg) {
        config.elements.prov3Msg.innerHTML = '';
    }	
    if (config.elements.pcMessage) {
        config.elements.pcMessage.innerHTML = '';
    }
}

async function verifyFields(ev, config) {
    ev.preventDefault();

    if (config.skipVerification) {
        return;
    }

    try {
        if (
            config.elements.line1.length == 0 ||
            !config.elements.line1[0].value
        ) {
            setErrorIfExists(
                {
                    line1: ['Missing Value: Line 1'],
                },
                config
            );
            return;
        }
        var result = await verify(config);
        for (let i = 0; i < config.elements.line1.length; i++) {
            setValueIfExists(result, 'line1', config.elements.line1[i]);
        }
        if (result.zipPlus4) {
            result.postalOrZip = result.postalOrZip
                ? result.postalOrZip + '-' + result.zipPlus4
                : result.zipPlus4;
        }
        if (result.urbanization) {
            result.line2 = result.line2
                ? result.line2 + ' URB ' + result.urbanization
                : result.urbanization;
        }
        setValueIfExists(result, 'line2', config.elements.line2);
        setValueIfExists(result, 'city', config.elements.city);
        setValueIfExists(result, 'provinceOrState', config.elements.prov);
		//TODO: 
        setValueIfExists(result, 'postalOrZip', config.elements.pc);
        setValueIfExists(result, 'country', config.elements.country);
        if (!config.isInternational) {
            setErrorIfExists(result.errors, config);
        }
        const comp = config.elements.status;
        if (isTextBox(comp)) {
            comp.value = config.isInternational
                ? result.summary.verificationStatus
                : result.status;
        } else {
            comp.innerHTML = config.isInternational
                ? result.summary.verificationStatus
                : result.status;
        }
        // If status is verified or corrected then submit the form
        if (result.status === 'verified' || result.status === 'corrected') {
            for (let i = 0; i < config.elements.invalidBox.length; i++) {
                if (config.elements.invalidBox[i].hasAttribute('required')) {
                    config.elements.invalidBox[i].removeAttribute('required');
                }
            }
            config.elements.form.submit();
        }
    } catch (err) {
        return console.warn(err);
    }
}

// *********************************************** Common Functions
// ***********************************************

function getElementValue(element) {
    if (isTextBox(element)) {
        return element.value;
    } else {
        return element.innerHTML;
    }
}

// Check if elem is textbox

function isTextBox(element) {
    var tagName = element.tagName.toLowerCase();
    if (tagName === 'textarea' || tagName === 'input' || tagName == 'select') {
        return true;
    }
    return false;
}

function setValueIfExists(option, fieldName, component) {
    if (component && option[fieldName]) {
        if (isTextBox(component)) {
			if((fieldName == "country" || fieldName == "countryCode") && option[fieldName] == "US"){
				component.value = "USA";
			}else if((fieldName == "country" || fieldName == "countryCode") && option[fieldName] == "CA"){
				component.value = "Canada";			
			}else {
				component.value = option[fieldName] || '';
			}	
        } else {
            component.innerHTML = option[fieldName] || '';
        }
    }
}

// ************************************************** Errors
// ******************************************************

function setErrorForArray(element, array) {
    if (!element) {
        return;
    }
    element.innerHTML = '';
    if (!Array.isArray(array)) {
        return;
    } else {
        var ul = document.createElement('ul');
        for (let i = 0; i < array.length; i++) {
            var li = document.createElement('li');
            var text = document.createTextNode(array[i]);
            li.appendChild(text);
            ul.appendChild(li);
        }
        element.appendChild(ul);
    }
}

function setErrorIfExists(errors, config) {
    if (errors.line1 && errors.line1.length > 0) {
        for (let i = 0; i < config.elements.line1Msg.length; i++) {
            setErrorForArray(config.elements.line1Msg[i], errors.line1);
        }
    } else {
        for (let i = 0; i < config.elements.line1Msg.length; i++) {
            setErrorForArray(config.elements.line1Msg[i], null);
        }
    }

    if (errors.generic && errors.generic.length > 0) {
        setErrorForArray(config.elements.errorBox, errors.generic);
    } else {
        setErrorForArray(config.elements.errorBox, null);
    }

    if (errors.city && errors.city.length > 0) {
        setErrorForArray(config.elements.cityMsg, errors.city);
    } else {
        setErrorForArray(config.elements.cityMsg, null);
    }

    if (errors.provinceOrState && errors.provinceOrState.length > 0) {
        setErrorForArray(config.elements.provMsg, errors.provinceOrState);
    } else {
        setErrorForArray(config.elements.provMsg, null);
    }

    if (errors.postalOrZip && errors.postalOrZip.length > 0) {
        setErrorForArray(config.elements.pcMessage, errors.postalOrZip);
    } else {
        setErrorForArray(config.elements.pcMessage, null);
    }
}

// ********************************* Load styles initially on document load
// ***************************************

var styles = '.pg-input-wrapper { position: relative; display: inline;}';
styles +=
    ' .pg-autocomplete-list { position: absolute; border: 1px solid #d4d4d4; border-bottom: none; border-top: none; z-index: 99; /*position the autocomplete items to be the same width as the container:*/  left: 0; right: 0; }';
styles +=
    ' .pg-autocomplete-list div { padding: 8px 10px; cursor: pointer; background-color: #fff; border-bottom: 1px solid #d4d4d4; }';
styles += ' .pg-autocomplete-list div:hover { background-color: #e9e9e9; }';
styles +=
    ' [data-pg-address-active] {background-color: DodgerBlue !important;color: #ffffff;}';

function addStyle(styles) {
    /* Create style element */
    var css = document.createElement('style');
    css.type = 'text/css';

    if (css.styleSheet) css.styleSheet.cssText = styles;
    else css.appendChild(document.createTextNode(styles));

    /* Append style to the head element */
    document.getElementsByTagName('head')[0].appendChild(css);

    // ClickFunnel disable browser autofill
    for (const div of addressDivs) {
        div.setAttribute('autocomplete', 'new-password');
    }
}
document.onload = addStyle(styles);

function registerPostGridAddressCompletedCallback(callback) {
    const maximum = Object.keys(customAddressCompletedCallbacks).reduce(
        (previousMaximum, currentValue) =>
            Math.max(previousMaximum, parseInt(currentValue)),
        0
    );
    const nextKey = maximum + 1;

    customAddressCompletedCallbacks[nextKey] = callback;
    return nextKey;
}

function unregisterPostGridAddressCompletedCallback(index) {
    delete customAddressCompletedCallbacks[index];
}