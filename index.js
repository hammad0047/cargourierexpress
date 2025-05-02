const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const received_awb = process.argv[2];
console.log(`Received AWB Number: ${received_awb}`);

const awb = received_awb
const awb2 = received_awb
const price_class = "VEGETABLE"
const pieces = "200"
const weight = "2500"
const volume = "3"
const commodity = "VEGETABLE"
const flight_no = "PK283"
const booking_date = "20May"
const departure_date = "30Apr"
const shipper_name = "ABDUL MOEZ TRADERS"
const shipper_street = "22KM FEROZPUR ROAD"
const shipper_city = "Peshawar"
const shipper_city_short = "PEW"
const shipper_country_short = "PK"
const shipper_country = "PAKISTAN"
const consignee_name = "NICE AL MARRI EXPRESS SER"
const consignee_street = "COURIER MAIL CENTER FG2 D"
const consignee_city = "DUBAI"
const consignee_city_short = "DXB"
const consignee_country_short = "AE"
const consignee_country = "United Arab Emirates"

const username = '';
const password = '';


const indexUrl = "https://cargospot-portal.champ.aero/index.asp?portal_id=PIA";
const newBookingUrl = `https://cargospot-portal.champ.aero/booking_confirm.asp?is_template_exist=no&template_type=FFR&save_template=no&save_draft=no&use_stock_control=0&use_range=0&flight_info=${encodeURIComponent(flight_no)}%2C${encodeURIComponent(booking_date)}%2C${encodeURIComponent(shipper_city_short)}%2C${encodeURIComponent(consignee_city_short)}%2C10%3A00%2C%2C15%3A15%2C%2C320%2C0%2CNEED+SPACE&iata_cass=41551395314&max_weight=5600&max_flight=5&book_type_text=booking&uom=M&selected_uom=M&shc_count=9&dim_ro=&dim_req=no&dimension_unit=&participant_type=Forwarder&is_valid_awb=yes&is_product_enabled=no&selected_product=&product_text=&api_id=csp&json_selected_uld_types=%7B%22selected_ulds%22%3A%5B%5D%7D&selected_airline=${encodeURIComponent(shipper_country_short)}&original_shipment_info=${encodeURIComponent(shipper_city)}+%28${encodeURIComponent(shipper_city_short)}%29+-+${encodeURIComponent(shipper_country)}+%5B${encodeURIComponent(shipper_country_short)}%5D${encodeURIComponent(consignee_city)}+%28${encodeURIComponent(consignee_city_short)}%29+-+${encodeURIComponent(consignee_country)}+%5B${encodeURIComponent(consignee_country_short)}%5D${encodeURIComponent(weight)}${encodeURIComponent(volume)}&err=&al_inter_date=&original_account_weight_valuation=Prepaid&selected_account_weight_valuation=Prepaid&selected_account_number=&calc_counter=0&prefix=214&awb=${encodeURIComponent(awb)}&booking_account_number=41551395314&price_class=${encodeURIComponent(price_class)}&shc_1=&shc_2=&shc_3=&shc_4=&shc_5=&shc_6=&shc_7=&shc_8=&shc_9=&origin=${encodeURIComponent(shipper_city)}+%28${encodeURIComponent(shipper_city_short)}%29+-+${encodeURIComponent(shipper_country)}+%5BPK%5D&pieces=${encodeURIComponent(pieces)}&destination=${encodeURIComponent(consignee_city)}+%28${encodeURIComponent(consignee_city_short)}%29+-+${encodeURIComponent(consignee_country)}+%5B${encodeURIComponent(consignee_country_short)}%5D&weight=${encodeURIComponent(weight)}&weight_code=KG&commodity=${encodeURIComponent(commodity)}&volume_info=&volume=${encodeURIComponent(volume)}&volume_code=MC&adhoc_rate_amount=&account_weight_valuation=Prepaid&account_other=Prepaid&special_handling_info_1=&special_handling_info_2=&uld_details=close&orig_uld_type_1=&orig_loaded_wgt_1=&orig_uld_unit_1=&orig_uld_volume_1=&orig_uld_volume_unit_1=&orig_uld_currency_1=&orig_charge_per_uld_1=&orig_over_pivot_rate_1=&orig_uld_id_1=&orig_adhoc_id_1=&uld_unit_1=KG&uld_status_1=&orig_uld_status_1=&uld_id_1=&adhoc_id_1=&orig_uld_type_2=&orig_loaded_wgt_2=&orig_uld_unit_2=&orig_uld_volume_2=&orig_uld_volume_unit_2=&orig_uld_currency_2=&orig_charge_per_uld_2=&orig_over_pivot_rate_2=&orig_uld_id_2=&orig_adhoc_id_2=&uld_unit_2=KG&uld_status_2=&orig_uld_status_2=&uld_id_2=&adhoc_id_2=&orig_uld_type_3=&orig_loaded_wgt_3=&orig_uld_unit_3=&orig_uld_volume_3=&orig_uld_volume_unit_3=&orig_uld_currency_3=&orig_charge_per_uld_3=&orig_over_pivot_rate_3=&orig_uld_id_3=&orig_adhoc_id_3=&uld_unit_3=KG&uld_status_3=&orig_uld_status_3=&uld_id_3=&adhoc_id_3=&orig_uld_type_4=&orig_loaded_wgt_4=&orig_uld_unit_4=&orig_uld_volume_4=&orig_uld_volume_unit_4=&orig_uld_currency_4=&orig_charge_per_uld_4=&orig_over_pivot_rate_4=&orig_uld_id_4=&orig_adhoc_id_4=&uld_unit_4=KG&uld_status_4=&orig_uld_status_4=&uld_id_4=&adhoc_id_4=&orig_uld_type_5=&orig_loaded_wgt_5=&orig_uld_unit_5=&orig_uld_volume_5=&orig_uld_volume_unit_5=&orig_uld_currency_5=&orig_charge_per_uld_5=&orig_over_pivot_rate_5=&orig_uld_id_5=&orig_adhoc_id_5=&uld_unit_5=KG&uld_status_5=&orig_uld_status_5=&uld_id_5=&adhoc_id_5=&shipper_consignee=expand&shipper_name=${encodeURIComponent(shipper_name)}&consignee_name=${encodeURIComponent(consignee_name)}&shipper_street=${encodeURIComponent(shipper_street)}&consignee_street=${encodeURIComponent(consignee_street)}&shipper_city=${encodeURIComponent(shipper_city)}&consignee_city=${encodeURIComponent(consignee_city)}&shipper_zipcode=&consignee_zipcode=&shipper_state=&consignee_state=&shipper_country=${encodeURIComponent(shipper_country_short)}&consignee_country=${encodeURIComponent(consignee_country_short)}&shipper_contact_type=&consignee_contact_type=&shipper_contact_number=&consignee_contact_number=&remark_1=&remark_2=&departure_date=${encodeURIComponent(departure_date)}&departure_time=&arrival_date=&submit_button=&template_name=&allot_ref=&submit_booking=none&submit_allotment=none&submit_interline=none&submit_manual=none`;
const newBookingUrlWithNewAwb = `https://cargospot-portal.champ.aero/booking_confirm.asp?is_template_exist=no&template_type=FFR&save_template=no&save_draft=no&use_stock_control=0&use_range=0&flight_info=${encodeURIComponent(flight_no)}%2C${encodeURIComponent(booking_date)}%2C${encodeURIComponent(shipper_city_short)}%2C${encodeURIComponent(consignee_city_short)}%2C10%3A00%2C%2C15%3A15%2C%2C320%2C0%2CNEED+SPACE&iata_cass=41551395314&max_weight=5600&max_flight=5&book_type_text=booking&uom=M&selected_uom=M&shc_count=9&dim_ro=&dim_req=no&dimension_unit=&participant_type=Forwarder&is_valid_awb=yes&is_product_enabled=no&selected_product=&product_text=&api_id=csp&json_selected_uld_types=%7B%22selected_ulds%22%3A%5B%5D%7D&selected_airline=${encodeURIComponent(shipper_country_short)}&original_shipment_info=${encodeURIComponent(shipper_city)}+%28${encodeURIComponent(shipper_city_short)}%29+-+${encodeURIComponent(shipper_country)}+%5B${encodeURIComponent(shipper_country_short)}%5D${encodeURIComponent(consignee_city)}+%28${encodeURIComponent(consignee_city_short)}%29+-+${encodeURIComponent(consignee_country)}+%5B${encodeURIComponent(consignee_country_short)}%5D${encodeURIComponent(weight)}${encodeURIComponent(volume)}&err=&al_inter_date=&original_account_weight_valuation=Prepaid&selected_account_weight_valuation=Prepaid&selected_account_number=&calc_counter=0&prefix=214&awb=${encodeURIComponent(awb2)}&booking_account_number=41551395314&price_class=${encodeURIComponent(price_class)}&shc_1=&shc_2=&shc_3=&shc_4=&shc_5=&shc_6=&shc_7=&shc_8=&shc_9=&origin=${encodeURIComponent(shipper_city)}+%28${encodeURIComponent(shipper_city_short)}%29+-+${encodeURIComponent(shipper_country)}+%5BPK%5D&pieces=${encodeURIComponent(pieces)}&destination=${encodeURIComponent(consignee_city)}+%28${encodeURIComponent(consignee_city_short)}%29+-+${encodeURIComponent(consignee_country)}+%5B${encodeURIComponent(consignee_country_short)}%5D&weight=${encodeURIComponent(weight)}&weight_code=KG&commodity=${encodeURIComponent(commodity)}&volume_info=&volume=${encodeURIComponent(volume)}&volume_code=MC&adhoc_rate_amount=&account_weight_valuation=Prepaid&account_other=Prepaid&special_handling_info_1=&special_handling_info_2=&uld_details=close&orig_uld_type_1=&orig_loaded_wgt_1=&orig_uld_unit_1=&orig_uld_volume_1=&orig_uld_volume_unit_1=&orig_uld_currency_1=&orig_charge_per_uld_1=&orig_over_pivot_rate_1=&orig_uld_id_1=&orig_adhoc_id_1=&uld_unit_1=KG&uld_status_1=&orig_uld_status_1=&uld_id_1=&adhoc_id_1=&orig_uld_type_2=&orig_loaded_wgt_2=&orig_uld_unit_2=&orig_uld_volume_2=&orig_uld_volume_unit_2=&orig_uld_currency_2=&orig_charge_per_uld_2=&orig_over_pivot_rate_2=&orig_uld_id_2=&orig_adhoc_id_2=&uld_unit_2=KG&uld_status_2=&orig_uld_status_2=&uld_id_2=&adhoc_id_2=&orig_uld_type_3=&orig_loaded_wgt_3=&orig_uld_unit_3=&orig_uld_volume_3=&orig_uld_volume_unit_3=&orig_uld_currency_3=&orig_charge_per_uld_3=&orig_over_pivot_rate_3=&orig_uld_id_3=&orig_adhoc_id_3=&uld_unit_3=KG&uld_status_3=&orig_uld_status_3=&uld_id_3=&adhoc_id_3=&orig_uld_type_4=&orig_loaded_wgt_4=&orig_uld_unit_4=&orig_uld_volume_4=&orig_uld_volume_unit_4=&orig_uld_currency_4=&orig_charge_per_uld_4=&orig_over_pivot_rate_4=&orig_uld_id_4=&orig_adhoc_id_4=&uld_unit_4=KG&uld_status_4=&orig_uld_status_4=&uld_id_4=&adhoc_id_4=&orig_uld_type_5=&orig_loaded_wgt_5=&orig_uld_unit_5=&orig_uld_volume_5=&orig_uld_volume_unit_5=&orig_uld_currency_5=&orig_charge_per_uld_5=&orig_over_pivot_rate_5=&orig_uld_id_5=&orig_adhoc_id_5=&uld_unit_5=KG&uld_status_5=&orig_uld_status_5=&uld_id_5=&adhoc_id_5=&shipper_consignee=expand&shipper_name=${encodeURIComponent(shipper_name)}&consignee_name=${encodeURIComponent(consignee_name)}&shipper_street=${encodeURIComponent(shipper_street)}&consignee_street=${encodeURIComponent(consignee_street)}&shipper_city=${encodeURIComponent(shipper_city)}&consignee_city=${encodeURIComponent(consignee_city)}&shipper_zipcode=&consignee_zipcode=&shipper_state=&consignee_state=&shipper_country=${encodeURIComponent(shipper_country_short)}&consignee_country=${encodeURIComponent(consignee_country_short)}&shipper_contact_type=&consignee_contact_type=&shipper_contact_number=&consignee_contact_number=&remark_1=&remark_2=&departure_date=${encodeURIComponent(departure_date)}&departure_time=&arrival_date=&submit_button=&template_name=&allot_ref=&submit_booking=none&submit_allotment=none&submit_interline=none&submit_manual=none`;
const update_BookingUrl = `https://cargospot-portal.champ.aero/booking_update_confirm.asp?original_shipment_info=Peshawar+%28PEW%29+-+Pakistan+%5BPK%5DDubai+%28DXB%29+-+United+Arab+Emirates+%5BAE%5D10001&old_flight=PK283%2C12Aug2024%2CPEW%2CDXB%2C14%3A25%2C%2C16%3A45%2C%2C32A%2C0%2CNeed+Space%2CCN&max_weight=9999999&max_flight=5&book_type=booking&new_book_type=booking&uom=M&selected_uom=M&check_awb=yes&volume_info=&dimension_unit=CM&shc_count=9&old_status=Cancelled&dim_req=false&save_draft=no&is_valid_awb=yes&err=&orig_origin=${encodeURIComponent(shipper_city_short)}&orig_origin_text=${encodeURIComponent(shipper_city)}+%28${encodeURIComponent(shipper_city_short)}%29+-+${encodeURIComponent(shipper_country)}+%5B${encodeURIComponent(shipper_country_short)}%5D&orig_destination=${encodeURIComponent(consignee_city_short)}&orig_destination_text=+%28${encodeURIComponent(consignee_city_short)}%29+-+${encodeURIComponent(consignee_country)}+%5B${encodeURIComponent(consignee_country_short)}%5D&orig_pieces=${encodeURIComponent(pieces)}&orig_weight=${encodeURIComponent(weight)}&orig_weight_code=KG&orig_volume=${encodeURIComponent(volume)}&orig_volume_code=MC&orig_volume_info=&orig_product=&orig_price_class=${encodeURIComponent(price_class)}&orig_handling_code=&orig_product_code=&orig_commodity=${encodeURIComponent(commodity)}&is_routing_complete=true&adhoc_id=&orig_adhoc_id=&adhoc_status=&orig_adhoc_status=&has_new_flight=0&find_new_flight=1&orig_adhoc_rate_currency=&orig_adhoc_rate_amount=&is_uld_rate_enabled=0&api_id=csp&is_product_enabled=no&selected_product=&price_class_session=${encodeURIComponent(price_class)}&selected_airline=${encodeURIComponent(flight_no)}&json_selected_uld_types=%7B%22selected_ulds%22%3A%5B%5D%7D&booking_id=10952027&flight_info=${encodeURIComponent(flight_no)}%2C${encodeURIComponent(booking_date)}%2C${encodeURIComponent(shipper_city_short)}%2C${encodeURIComponent(consignee_city_short)}%2C15%3A40%2C%2C18%3A00%2C%2C320%2C0%2CNEED+SPACE&is_new_flight_different=0&etag=W%2F%221721304470192%22&adhoc_uld_impact_warning=0&original_account_weight_valuation=Prepaid&selected_account_weight_valuation=Prepaid&booking_account_number=27301040012&calc_counter=0&prefix=214&awb=${encodeURIComponent(awb)}&price_class=${encodeURIComponent(price_class)}&shc_1=&shc_2=&shc_3=&shc_4=&shc_5=&shc_6=&shc_7=&shc_8=&shc_9=&origin=${encodeURIComponent(shipper_city)}+%28${encodeURIComponent(shipper_city_short)}%29+-+${encodeURIComponent(shipper_country)}+%5B${encodeURIComponent(shipper_country_short)}%5D&pieces=${encodeURIComponent(pieces)}&destination=${encodeURIComponent(consignee_city)}+%28${encodeURIComponent(consignee_city_short)}%29+-+${encodeURIComponent(consignee_country)}+%5B${encodeURIComponent(consignee_country_short)}%5D&weight=${encodeURIComponent(weight)}&weight_code=KG&commodity=${encodeURIComponent(commodity)}&volume=${encodeURIComponent(volume)}&volume_code=MC&adhoc_rate_amount=&account_weight_valuation=Prepaid&account_other=Prepaid&special_handling_info_1=&special_handling_info_2=&uld_details=close&orig_uld_type_1=&orig_loaded_wgt_1=&orig_uld_unit_1=&orig_uld_volume_1=&orig_uld_volume_unit_1=&orig_uld_currency_1=&orig_charge_per_uld_1=&orig_over_pivot_rate_1=&orig_uld_id_1=&orig_adhoc_id_1=&uld_unit_1=KG&uld_status_1=&orig_uld_status_1=&uld_id_1=&adhoc_id_1=&orig_uld_type_2=&orig_loaded_wgt_2=&orig_uld_unit_2=&orig_uld_volume_2=&orig_uld_volume_unit_2=&orig_uld_currency_2=&orig_charge_per_uld_2=&orig_over_pivot_rate_2=&orig_uld_id_2=&orig_adhoc_id_2=&uld_unit_2=KG&uld_status_2=&orig_uld_status_2=&uld_id_2=&adhoc_id_2=&orig_uld_type_3=&orig_loaded_wgt_3=&orig_uld_unit_3=&orig_uld_volume_3=&orig_uld_volume_unit_3=&orig_uld_currency_3=&orig_charge_per_uld_3=&orig_over_pivot_rate_3=&orig_uld_id_3=&orig_adhoc_id_3=&uld_unit_3=KG&uld_status_3=&orig_uld_status_3=&uld_id_3=&adhoc_id_3=&orig_uld_type_4=&orig_loaded_wgt_4=&orig_uld_unit_4=&orig_uld_volume_4=&orig_uld_volume_unit_4=&orig_uld_currency_4=&orig_charge_per_uld_4=&orig_over_pivot_rate_4=&orig_uld_id_4=&orig_adhoc_id_4=&uld_unit_4=KG&uld_status_4=&orig_uld_status_4=&uld_id_4=&adhoc_id_4=&orig_uld_type_5=&orig_loaded_wgt_5=&orig_uld_unit_5=&orig_uld_volume_5=&orig_uld_volume_unit_5=&orig_uld_currency_5=&orig_charge_per_uld_5=&orig_over_pivot_rate_5=&orig_uld_id_5=&orig_adhoc_id_5=&uld_unit_5=KG&uld_status_5=&orig_uld_status_5=&uld_id_5=&adhoc_id_5=&shipper_consignee=expand&shipper_name=${encodeURIComponent(shipper_name)}&consignee_name=${encodeURIComponent(consignee_name)}&shipper_street=${encodeURIComponent(shipper_street)}&consignee_street=${encodeURIComponent(consignee_street)}&shipper_city=${encodeURIComponent(shipper_city)}&consignee_city=${encodeURIComponent(consignee_city)}&shipper_zipcode=&consignee_zipcode=&shipper_state=&consignee_state=&shipper_country=${encodeURIComponent(shipper_country_short)}&consignee_country=${encodeURIComponent(consignee_country_short)}&shipper_contact_type=&consignee_contact_type=&shipper_contact_number=&consignee_contact_number=&remark_1=&remark_2=&orig_first_depature_date=16Jul&departure_date=${encodeURIComponent(departure_date)}&departure_time=&arrival_date=&allot_ref=&submit_allotment=none&submit_interline=none&submit_manual=none&submit_button=`;
//console.log(newBookingUrl);
(async () => {
    const startTime = Date.now();
    const timings = {
        browserLaunch: 0,
        login: 0,
        booking: 0,
        total: 0
    };

    try {
        console.log('Starting browser launch...');
        const browserStartTime = Date.now();
        const browser = await puppeteer.launch({
            headless: true,
            executablePath: '/usr/bin/google-chrome',
        });
        timings.browserLaunch = Date.now() - browserStartTime;
        console.log(`Browser launched in ${(timings.browserLaunch/1000).toFixed(2)}s`);

        const page = await browser.newPage();
        
        console.log('Navigating to index page...');
        const indexPageStartTime = Date.now();
        try {
            await page.goto(indexUrl, {waitUntil: 'networkidle2'});
        } catch (error) {
            console.error("Failed to load indexUrl:", error);
            await browser.close();
            return;
        }
        const indexPageTime = Date.now() - indexPageStartTime;
        console.log(`Index page loaded in ${(indexPageTime/1000).toFixed(2)}s`);

        console.log('Performing login...');
        const loginStartTime = Date.now();
        await page.type('.username', username);
        await page.type('.password', password);
        await page.click('#btn-login');

        try {
            await page.waitForNavigation({waitUntil: 'networkidle2'});
            timings.login = Date.now() - loginStartTime;
            console.log(`Login completed in ${(timings.login/1000).toFixed(2)}s`);
        } catch (error) {
            console.error("Login navigation failed:", error);
            await browser.close();
            return;
        }

        console.log('Navigating to booking page...');
        const bookingPageStartTime = Date.now();
        try {
            await page.goto(newBookingUrl, {waitUntil: 'networkidle2'});
        } catch (error) {
            console.error("Failed to load newBookingUrl:", error);
            await browser.close();
            return;
        }
        timings.booking = Date.now() - bookingPageStartTime;
        console.log(`Booking page loaded in ${(timings.booking/1000).toFixed(2)}s`);

        let currentUrl = page.url();
        console.log('Current Url', currentUrl);

        if (currentUrl.includes("booking_confirm.asp")) {
            try {
                console.log('Clicking on Confirm Button...');
                const confirmStartTime = Date.now();
                await page.click('#confirm');
                const confirmTime = Date.now() - confirmStartTime;
                console.log(`Confirm button clicked in ${(confirmTime/1000).toFixed(2)}s`);
                await page.waitForNavigation({waitUntil: 'networkidle2'});
                currentUrl = page.url();
                if (currentUrl.includes("create_booking_bridge.asp") || currentUrl.includes("booking_complete.asp")) {
                    console.log('Exiting browser.');
                    await browser.close();
                    return;
                }
            } catch (error) {
                console.error("Error during booking confirmation:", error);
            }
        }

        if (currentUrl.includes("booking.asp")) {
            let messageText = "";
            try {
                const messageElement = await page.$('#message');
                if (messageElement) {
                    messageText = await page.evaluate(el => el.textContent.trim(), messageElement);
                    console.log("Portal Message:", messageText);
                } else {
                    console.log("Message element not found.");
                }
            } catch (error) {
                console.error("Error extracting message text:", error);
            }

            if (messageText.includes("Already")) {
                console.log("Already booked, exiting browser.");
                //await browser.close();
                //return;
                console.log('\nPerformance Summary:');
                console.log('-------------------');
                console.log(`Browser Launch: ${timings.browserLaunch.toFixed(2)}ms (${(timings.browserLaunch/1000).toFixed(2)}s)`);
                console.log(`Login Process: ${timings.login.toFixed(2)}ms (${(timings.login/1000).toFixed(2)}s)`);
                console.log(`Booking Process: ${timings.booking.toFixed(2)}ms (${(timings.booking/1000).toFixed(2)}s)`);
                console.log(`Total Execution Time: ${timings.total.toFixed(2)}ms (${(timings.total/1000).toFixed(2)}s)`);
                console.log('-------------------\n');
                const excelData = {
                    timestamp: new Date().toISOString(),
                    awb: received_awb,
                    browserLaunch_ms: timings.browserLaunch.toFixed(2),
                    browserLaunch_s: (timings.browserLaunch/1000).toFixed(2),
                    login_ms: timings.login.toFixed(2),
                    login_s: (timings.login/1000).toFixed(2),
                    booking_ms: timings.booking.toFixed(2),
                    booking_s: (timings.booking/1000).toFixed(2),
                    total_ms: timings.total.toFixed(2),
                    total_s: (timings.total/1000).toFixed(2)
                };
        
                const excelFilePath = path.join(__dirname, 'performance_log.xlsx');
                let existingData = [];
        
                try {
                    if (fs.existsSync(excelFilePath)) {
                        const workbook = XLSX.readFile(excelFilePath);
                        existingData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                    }

                    existingData.push(excelData);
                    const worksheet = XLSX.utils.json_to_sheet(existingData);
                    const workbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(workbook, worksheet, 'Performance Log');
                    XLSX.writeFile(workbook, excelFilePath);
                    console.log(`Performance data saved to ${excelFilePath}`);
                } catch (excelError) {
                    console.error('Error saving to Excel:', excelError);
                }
            }
            if (messageText.includes("Not Found")) {
                console.log("Flight Not Found, Finding the flight again.");
                await page.goto(newBookingUrlWithNewAwb, {waitUntil: 'networkidle2'});
                await page.click('#confirm');
                console.log('CONFIRM_CLICKED'); // <-- This is your signal to trigger the next AWB
                await page.waitForNavigation({waitUntil: 'networkidle2'});
                currentUrl = page.url();
                if (currentUrl.includes("create_booking_bridge.asp") || currentUrl.includes("booking_complete.asp")) {
                    console.log('Exiting browser.');
                    await browser.close();
                    return;
                }
            }

            try {
                console.log("Retrying booking with new AWB.");
                //const newPage = await browser.newPage();
                await page.goto(newBookingUrlWithNewAwb, {waitUntil: 'networkidle2'});
                await page.click('#confirm');
                console.log('CONFIRM_CLICKED'); // <-- This is your signal to trigger the next AWB
                await page.waitForNavigation({waitUntil: 'networkidle2'});
                currentUrl = page.url();
                if (currentUrl.includes("create_booking_bridge.asp") || currentUrl.includes("booking_complete.asp")) {
                    console.log('Exiting browser.');
                    await browser.close();
                    return;
                }
            } catch (error) {
                console.log("Again Retrying booking with new AWB.");
                const newPage = await browser.newPage();
                await newPage.goto(newBookingUrl, {waitUntil: 'networkidle2'});
                await newPage.click('#confirm');
                console.log('CONFIRM_CLICKED'); // <-- This is your signal to trigger the next AWB
                await newPage.waitForNavigation({waitUntil: 'networkidle2'});
                currentUrl = newPage.url();
                if (currentUrl.includes("create_booking_bridge.asp") || currentUrl.includes("booking_complete.asp")) {
                    console.log('Exiting browser.');
                    await browser.close();
                    return;
                }
            }
        }

        timings.total = Date.now() - startTime;

        // Print performance summary
        console.log('\nPerformance Summary:');
        console.log('-------------------');
        console.log(`Browser Launch: ${timings.browserLaunch.toFixed(2)}ms (${(timings.browserLaunch/1000).toFixed(2)}s)`);
        console.log(`Login Process: ${timings.login.toFixed(2)}ms (${(timings.login/1000).toFixed(2)}s)`);
        console.log(`Booking Process: ${timings.booking.toFixed(2)}ms (${(timings.booking/1000).toFixed(2)}s)`);
        console.log(`Total Execution Time: ${timings.total.toFixed(2)}ms (${(timings.total/1000).toFixed(2)}s)`);
        console.log('-------------------\n');

        // Save to Excel
        const excelData = {
            timestamp: new Date().toISOString(),
            awb: received_awb,
            browserLaunch_ms: timings.browserLaunch.toFixed(2),
            browserLaunch_s: (timings.browserLaunch/1000).toFixed(2),
            login_ms: timings.login.toFixed(2),
            login_s: (timings.login/1000).toFixed(2),
            booking_ms: timings.booking.toFixed(2),
            booking_s: (timings.booking/1000).toFixed(2),
            total_ms: timings.total.toFixed(2),
            total_s: (timings.total/1000).toFixed(2)
        };

        const excelFilePath = path.join(__dirname, 'performance_log.xlsx');
        let existingData = [];

        try {
            if (fs.existsSync(excelFilePath)) {
                const workbook = XLSX.readFile(excelFilePath);
                existingData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            }

            existingData.push(excelData);
            const worksheet = XLSX.utils.json_to_sheet(existingData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Performance Log');
            XLSX.writeFile(workbook, excelFilePath);
            console.log(`Performance data saved to ${excelFilePath}`);
        } catch (excelError) {
            console.error('Error saving to Excel:', excelError);
        }
        
    } catch (globalError) {
        console.error("Unexpected script error:", globalError);
        timings.total = Date.now() - startTime;
        
        // Print performance summary even in case of error
        console.log('\nPerformance Summary (with error):');
        console.log('-------------------');
        console.log(`Browser Launch: ${timings.browserLaunch.toFixed(2)}ms (${(timings.browserLaunch/1000).toFixed(2)}s)`);
        console.log(`Login Process: ${timings.login.toFixed(2)}ms (${(timings.login/1000).toFixed(2)}s)`);
        console.log(`Booking Process: ${timings.booking.toFixed(2)}ms (${(timings.booking/1000).toFixed(2)}s)`);
        console.log(`Total Execution Time: ${timings.total.toFixed(2)}ms (${(timings.total/1000).toFixed(2)}s)`);
        console.log('-------------------\n');

        // Save error case to Excel
        const excelData = {
            timestamp: new Date().toISOString(),
            awb: received_awb,
            browserLaunch_ms: timings.browserLaunch.toFixed(2),
            browserLaunch_s: (timings.browserLaunch/1000).toFixed(2),
            login_ms: timings.login.toFixed(2),
            login_s: (timings.login/1000).toFixed(2),
            booking_ms: timings.booking.toFixed(2),
            booking_s: (timings.booking/1000).toFixed(2),
            total_ms: timings.total.toFixed(2),
            total_s: (timings.total/1000).toFixed(2),
            error: globalError.message
        };

        const excelFilePath = path.join(__dirname, 'performance_log.xlsx');
        let existingData = [];

        try {
            if (fs.existsSync(excelFilePath)) {
                const workbook = XLSX.readFile(excelFilePath);
                existingData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            }

            existingData.push(excelData);
            const worksheet = XLSX.utils.json_to_sheet(existingData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Performance Log');
            XLSX.writeFile(workbook, excelFilePath);
            console.log(`Performance data saved to ${excelFilePath}`);
        } catch (excelError) {
            console.error('Error saving to Excel:', excelError);
        }
    }
})();

