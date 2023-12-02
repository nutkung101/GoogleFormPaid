let templateSlideId = "1LKgN27OOhOAjBiGLBTOeBjvfe4B291Z-SDDP2lDl8Wk";
let folderResponsePdfId = "1kO_NTNG8QF2WN6_Svd4623kuE_hWch4R";
let folderResponseSlideId = "1EriJCyDntSsxxBU6ju0rq7lPOPBEyDke";
let sheetName = 'การตอบแบบฟอร์ม 1';
let pdf_file_name = "เอกสาร - ";
let isSendEmail = true;
let isSendLine = true;
let email_send_default = ['nattysung101@gmail.com'];
// let email_send_default = ['brilliantpy.live@gmail.com'];
// let email_send_default = ['brilliantpy1.live@gmail.com','brilliantpy2.live@gmail.com'];
var email_subject = 'หนังสือสัญญาประนีประนอมยอมความ';
var email_message = 'ข้อมูลส่งเข้าอีเมล์แล้ว';

let index_col = {"ประทับเวลา":0,"อีเมล์":1,"เลขที่รับแจ้ง":2,"วันที่เขียน":3,"ชื่อผู้ขับขี่รถประกัน":4,"ทะเบียนรถประกัน":5,"วันที่เกิดเหตุ":6,"สถานที่เกิดเหตุ":7,"เพื่อให้ทำสัญญา":8,"ทะเบียนคู่กรณี":9,"ค่าซ่อมรถคู่กรณี":10,"จำนวนวันอนุมติ":11,"วันละ":12,"รวมค่าขาดฯ":13,"รายการทรัพย์สิน":14,"ยุติค่าซ่อมจำนวน":15,"ชื่อบุคคล":16,"ยุติสินไหมจำนวน":17,"ชื่อ-สกุลคู่กรณี":18,"อายุ":19,"เลขบัตร":20,"วันออกบัตร":21,"วันหมดอายุ":22,"เชื้อชาติ":23,"สัญชาติ":24,"บ้านเลขที่":25,"หมู่ที่":26,"ชื่อซอย":27,"ชื่อถนน":28,"เขตตำบล":29,"เขตอำเภอ":30,"เขตจังหวัด":31,"เบอร์โทรติดต่อ":32,"การรับมอบอำนาจ":33,"รับมอบจาก":34,"เลขที่":35,"ถนน":36,"หมู่":37,"ตำบล":38,"อำเภอ":39,"จังหวัด":40,"เบอร์โทรผู้มอบอำนาจ":41,"โอนเงินเข้าบัญชี":42,"บัญชีธนาคาร":43,"แบบบัญชี":44,"ออมทรัพย์สาขา":45,"ออมทรัพย์เลขที่":46,"สาขากระแสรายวัน":47,"กระแสรายวันเลขที่":48,"เพื่อจ่ายค่า":49,"รวมเป็นเงิน":50,"อ่านว่า":51,"มีรายการเพิ่มหรือไม่":52,"ใส่ลายเซ็น":53,"ตราประทับบริษัท":54,"เพื่อให้ทำสัญญา[เพื่อจ่ายค่าซ่อมรถคู่กรณี]":55,"เพื่อให้ทำสัญญา[เพื่อจ่ายค่าขาดประโยชน์]":56,"เพื่อให้ทำสัญญา[เพื่อซ่อมทรัพย์สิน]":57,"เพื่อให้ทำสัญญา[เพื่อจ่ายสินไหมทดแทน]":58,"แบบบัญชี[ออมทรัพย์]":59,"แบบบัญชี[กระแสรายวัน]":60,"send_email_status":61,"send_line_status":62,"create_pdf_status":63};

let colEmail = index_col["อีเมล์"] || "";
// let colName = index_col["เลขเคลม"] || "";
let colEmailStatus = index_col["send_email_status"] || "";
let colLineStatus = index_col["send_line_status"] || "";
let colName = index_col["ชื่อ-สกุลคู่กรณี"] || "";
let colPdfStatus = index_col["create_pdf_status"] || "";
let colEmailStatusName = "BJ";
let colLineStatusName = "BK";
let colPdfStatusName = "BL";

let colAllImage = [
  { [index_col["ใส่ลายเซ็น"]]: "{{ใส่ลายเซ็น}}" },
  { [index_col["ตราประทับบริษัท"]]: "{{ตราประทับบริษัท}}" }, 
];
let index_col_checkbox = [
  { [index_col["แบบบัญชี"]]: [{ ออมทรัพย์: "BH" }, { กระแสรายวัน: "BI" }] },
  { [index_col["เพื่อให้ทำสัญญา"]]: [{ เพื่อจ่ายค่าซ่อมรถคู่กรณี: "BD" }, { เพื่อจ่ายค่าขาดประโยชน์: "BE" }, { เพื่อซ่อมทรัพย์สิน: "BF" }, { เพื่อจ่ายสินไหมทดแทน: "BG" }] },
];

let tokensV2 = ["7hvsGAuktSBNVrAauQTKSXE5MNe7NrgYsenDj1RraoM"]; // BrilliantPy line group(test only)
/*#########################  Editable1 End  #########################*/

// Init
let newSlideName = "New_FormToSlidePDF_";
let sent_status = "SENT";
let ss, sheet, lastRow, lastCol, range, values;
let data_name;
let newSlide, newSlideId, presentation, all_shape;
let titleName;
let exportPdf, pdf_name_full;
let email_send = [];
let filePath;




function formToSlidePdfLine() {
  /*######################### Editable2 Start #########################*/
  function formatMsgToLine() {
    return `[สัญญาประนีประนอม]
    

${filePath}`;
  }
  /*#########################  Editable2 End  #########################*/
  initSpreadSheet().then(async function () {
    formatTitle();
    for (let i = 1; i < lastRow; i++) {
      let numRow = i + 1;
      clearVal();
      let cur_data = values[i];
      data_name = cur_data[colName];
      try {
        data_name = data_name.replace(/\s/g, "");
      } catch (e) {}
      let emailStatus = cur_data[colEmailStatus];
      let lineStatus = cur_data[colLineStatus];
      if (
        (!isSendEmail || (isSendEmail && emailStatus == sent_status)) &&
        (!isSendLine || (isSendLine && lineStatus == sent_status))
      ) {
        continue;
      }
      await duplicateSlide().then(async function () {
        await updateCheckboxCol(cur_data, numRow).then(async function () {
          values = range.getValues();
          cur_data = values[i];
          try {
            cur_data[index_col["วันที่เขียน"]] = customFormatDate(
              cur_data[index_col["วันที่เขียน"]],
              "date",
              "dd/MM/yyyy"
            );
             cur_data[index_col["วันที่เกิดเหตุ"]] = customFormatDate(
              cur_data[index_col["วันที่เกิดเหตุ"]],
              "date",
              "dd/MM/yyyy"
            );
             cur_data[index_col["วันออกบัตร"]] = customFormatDate(
              cur_data[index_col["วันออกบัตร"]],
              "date",
              "dd/MM/yyyy"
            );
             cur_data[index_col["วันหมดอายุ"]] = customFormatDate(
              cur_data[index_col["วันหมดอายุ"]],
              "date",
              "dd/MM/yyyy"
            );
           
           
          
          } catch (e) {}
          await updateSlideData(cur_data).then(async function () {
            presentation.saveAndClose();
            await createPdf().then(async function () {
              removeTempSlide();
              let cur_email = cur_data[colEmail];
              let emailStatus = cur_data[colEmailStatus];
              let lineStatus = cur_data[colLineStatus];
              if (validateEmail(cur_email)) {
                email_send.push(cur_email);
              }
              console.log(email_send);
              if (isSendEmail && emailStatus != sent_status) {
                for (let j = 0; j < email_send.length; j++) {
                  if (validateEmail(email_send[j])) {
                    await sendEmailWithAttachment(email_send[j]).then(
                      function () {
                        if (j == email_send.length - 1) {
                          updateStatusSent(numRow, "email");
                        }
                      }
                    );
                  }
                }
              }
              if (isSendLine && lineStatus != sent_status) {
                sendLineNotify(formatMsgToLine());
                updateStatusSent(numRow, "line");
              }
            });
          });
        });
      });
    }
    console.log("Program completed");
  });
}

async function sendLineNotify(all_message_send) {
  return new Promise(function (resolve) {
    for (let k = 0; k < tokensV2.length; k++) {
      let formData = {
        message: all_message_send,
      };
      let options = {
        method: "post",
        payload: formData,
        headers: { Authorization: "Bearer " + tokensV2[k] },
      };
      UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
    }

    resolve();
    console.log("sendLineNotify completed");
  });
}

function clearVal() {
  data_name = "";
  newSlide = newSlideId = presentation = "";
  exportPdf = pdf_name_full = "";
  email_send = isSendEmail ? [...email_send_default] : [];
  all_shape = "";
  console.log("clearVal completed");
}

async function initSpreadSheet() {
  return new Promise(function (resolve) {
    ss = SpreadsheetApp.getActive();
    sheet = ss.getSheetByName(sheetName);
    lastRow = sheet.getLastRow();
    lastCol = sheet.getLastColumn();
    range = sheet.getDataRange();
    values = range.getValues();
    resolve();
    console.log("initSpreadSheet completed");
  });
}

function formatTitle() {
  titleName = values[0];
  titleName.forEach(function (item, index) {
    titleName[index] = "{{" + item + "}}";
  });
  console.log("formatTitle completed");
}

async function duplicateSlide() {
  return new Promise(function (resolve) {
    let templateSlide = DriveApp.getFileById(templateSlideId);
    let templateResponseFolder = DriveApp.getFolderById(folderResponseSlideId);
    newSlide = templateSlide.makeCopy(
      newSlideName.concat(data_name),
      templateResponseFolder
    );
    resolve();
    console.log("duplicateSlide completed");
  });
}

async function updateCheckboxCol(cur_data, numRow) {
  return new Promise(function (resolve) {
    index_col_checkbox.forEach(function (item) {
      Object.keys(item).forEach(function (key) {
        var cur_checkbox_val = cur_data[key];
        item[key].forEach(function (item_ele) {
          Object.keys(item_ele).forEach(function (key_item_ele) {
            if (key_item_ele === cur_checkbox_val) {
              sheet
                .getRange(item_ele[key_item_ele].concat(numRow))
                .setValue("✓");
            }
          });
        });
      });
    });
    //index_col_multi_checkbox.forEach(function (item) {

    resolve();
    console.log("updateCheckboxCol completed");
  });
}

async function updateSlideData(cur_data) {
  return new Promise(function (resolve) {
    // Init
    newSlideId = newSlide.getId();
    presentation = SlidesApp.openById(newSlideId);
    let slide = presentation.getSlides()[0];
    all_shape = slide.getShapes();
    titleName.forEach(async function (item, index) {
      let isColImg = false;
      colAllImage.forEach(async function (img_item) {
        Object.keys(img_item).forEach(async function (key) {
          if (item === img_item[key]) {
            all_shape.forEach(async function (s) {
              if (s.getText().asString().includes(img_item[key])) {
                let cur_img_url = cur_data[key];
                let imageFileId = getIdFromUrl(cur_img_url);
                if (imageFileId) {
                  isColImg = true;
                  let image = DriveApp.getFileById(imageFileId).getBlob();
                  await replaceImage(s, image).then(async function () {
                    console.log("replaceImage completed");
                  });
                }
              }
            });
          }
        });
      });
      if (!isColImg) {
        let templateVariable = item;
        let replaceValue = cur_data[index];
        presentation.replaceAllText(templateVariable, replaceValue);
      }
    });
    resolve();
    console.log("updateSlideData completed");
  });
}

async function replaceImage(s, image) {
  let res;
  return new Promise(function (resolve) {
    try {
      res = s.replaceWithImage(image);
    } catch (e) {
      console.log("error:", e);
    }
    if (res) {
      console.log("resolve");
      resolve();
    }
  });
}
async function createPdf() {
  return new Promise(function (resolve, reject) {
    let pdf = DriveApp.getFileById(newSlideId)
      .getBlob()
      .getAs("application/pdf");
    pdf_name_full = pdf_file_name + data_name + ".pdf";
    pdf.setName(pdf_name_full);
    exportPdf = DriveApp.getFolderById(folderResponsePdfId).createFile(pdf);
    filePath = exportPdf.getUrl();
    if (exportPdf) {
      resolve();
      console.log("สร้างไฟล์ PDF เสร็จสิ้น");
    } else {
      reject();
      console.log("สร้างไฟล์ PDF ได้");
    }
  });
}

async function sendEmailWithAttachment(email) {
  return new Promise(function (resolve, reject) {
    let file =
      DriveApp.getFolderById(folderResponsePdfId).getFilesByName(pdf_name_full);
    if (!file.hasNext()) {
      console.error("Could not open file " + pdf_name_full);
      return;
    }
    try {
      MailApp.sendEmail({
        to: email,
        subject: email_subject,
        htmlBody: email_message,
        attachments: [file.next().getAs(MimeType.PDF)],
      });
      resolve();
      console.log("sendEmailWithAttachment completed");
    } catch (e) {
      reject();
      console.log(
        "sendEmailWithAttachment error with email (" + email + "). " + e
      );
    }
  });
}

function removeTempSlide() {
  try {
    DriveApp.getFileById(newSlideId).setTrashed(true);
    console.log("removeTempSlide completed");
  } catch (e) {
    console.log("removeTempSlide error");
  }
}

function updateStatusSent(numRow, mode) {
  if (mode == "email") {
    sheet.getRange(colEmailStatusName.concat(numRow)).setValue(sent_status);
  } else if (mode == "line") {
    sheet.getRange(colLineStatusName.concat(numRow)).setValue(sent_status);
  } else if (mode == "both") {
    sheet.getRange(colEmailStatusName.concat(numRow)).setValue(sent_status);
    sheet.getRange(colLineStatusName.concat(numRow)).setValue(sent_status);
  }
  console.log("updateStatusSent completed");
}

function formatUrlImg(url) {
  let new_url = "";
  let start_url = "https://drive.google.com/uc?id=";
  new_url = start_url + getIdFromUrl(url);
  return new_url;
}

function getIdFromUrl(url) {
  return url.match(/[-\w]{25,}/);
}

function validateEmail(email) {
  var re = /\S+@\S+\.\S+/;
  if (!re.test(email)) {
    return false;
  } else {
    return true;
  }
}

function generateTitle() {
  let result = "";
  initSpreadSheet();
  let title = values[0];
  console.log("title:", title);
  for (let i = 0; i < title.length; i++) {
    result += `"${title[i]}":${i}`;
    if (i != title.length - 1) {
      result += ",";
    } else {
      result = `let index_col = {${result}};`;
    }
  }
  console.log(result);
}
function customFormatDate(date, mode, format) {
  let _timezone = "";
  if (mode == "date") {
    _timezone = "GMT+7";
  } else if (mode == "time") {
    _timezone = "GMT+6:43";
  } else {
    _timezone = timezone;
  }
  return Utilities.formatDate(date, _timezone, format);
}