function myFunction(e) {
  var sheet = SpreadsheetApp.getActiveSheet();
var row =  SpreadsheetApp.getActiveSheet().getLastRow();
  var guestName = e.values[4];
  var guestEmail = e.values[7];
  var test = 'test';
  var emailadmission = "admission@mutiaraharapan.sch.id"
  var Subject = "Greetings from Mutiara Harapan Islamic School";
  var body = "Assalamu’alaykum warahmatullahi wabarakatuh,<br><br>Dear Mr/Mrs. " + guestName + " , thank you for your interest in our school. <br><br>We would like to inform you that enrolment for the 2022/2023 school year is now open. Please visit the following link for access to our enrolment form: <br><br>https://bit.ly/MHISEnrolmentForm <br><br>We also provide the opportunity for any parents who wish to learn more about the school's programs to have an Interactive Call with our Principals through voice call or video call at the convenience of your available time.  Here we share with you the link to arrange Interactive Call :<br><br>https://bit.ly/InteractiveCall-MHIS <br><br>Should you have any questions or need us to assist you in the enrolment process, you may contact us through +62 813 8908 1220 or +62 821 2546 9320 at office hours from 8 a.m - 3 p.m.<br><br>Thank you for your kind attention. <br><br>Wassalamu’alaykum warahmatullahi wabarakatuh<br><br>Warm regards, <br>Mutiara Harapan Islamic School";
 var country = e.values[8];
var currentzip = e.values[9];
var Geocodingcurrent = UrlFetchApp.fetch("https://maps.googleapis.com/maps/api/geocode/json?address=" + currentzip +"+"+ country +"&key=AIzaSyDW0XaRzzDDMgUdPOXvPYrKej__-b6Cby4");
var jsoncurrent = Geocodingcurrent.getContentText();
  var datacurrent = JSON.parse(jsoncurrent);
  var filtered_kelurahan = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_4");
      })
  var filtered_kecamatan = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_3");
      })
  var filtered_kota = datacurrent.results[0].address_components.filter(function(address_component){
    return address_component.types.includes("administrative_area_level_2");
      })
  var kota = filtered_kota.length ? filtered_kota[0].long_name: "";
  var kecamatan = filtered_kecamatan.length ? filtered_kecamatan[0].long_name: "";
  var kelurahan = filtered_kelurahan.length ? filtered_kelurahan[0].long_name: "";
  var setkelurahan = sheet.getRange(row,20).setValue(kelurahan);
  var setkecamatan = sheet.getRange(row,19).setValue(kecamatan);
  var setkota = sheet.getRange(row,18).setValue(kota); 
  MailApp.sendEmail({
    to: (guestEmail + "," + emailadmission),
    subject: Subject,
    htmlBody: body,
  }); 
}
