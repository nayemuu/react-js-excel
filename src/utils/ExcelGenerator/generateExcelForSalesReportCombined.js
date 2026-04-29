import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { isValidNumber } from "../dataTypeCheck";
import { convertDate, convertTime } from "../ConvertDateTime";
import { roundTo } from "../utils";

const fromDateInLastQuery = "22-04-2026";
const toDateInLastQuery = "26-04-2026";

const dailyCombineSales = {
  data: {
    products_data: {
      results: [
        {
          invoice_number: "INV220426000001",
          order_source: "E-commerce",
          total_amount: "50.00",
          discount_amount: 0.5,
          vat_amount: "0.00",
          exchange_amount: "0.00",
          net_amount: "50.00",
          payment_method: "Cash On Delivery",
          paid_amount: "0.00",
          total_mrp: "50.00",
          profit_amount: "5.98",
          gp_amount: "11.95",
          shipping_charge: "0.00",
          created_time: "2026-04-22T11:07:00+06:00",
          redeem_point: 0,
          less_amount: "0.50",
          order_cost: "44.02",
          exchange_cost: "0.00",
          due: "50.00",
          cash_paid: "0.00",
          card_paid: "0.00",
          bkash_paid: "0.00",
          nagad_paid: "0.00",
          upay_paid: "0.00",
          city_paid: "0.00",
          nexus_paid: "0.00",
          rocket_paid: "0.00",
          ecom_cash_paid: "0.00",
          ecom_online_paid: "0.00",
          reference_invoice: null,
        },
        {
          invoice_number: "INV220426000003",
          order_source: "E-commerce",
          total_amount: "38.00",
          discount_amount: 0.6,
          vat_amount: "0.00",
          exchange_amount: "0.00",
          net_amount: "38.00",
          payment_method: "Cash On Delivery",
          paid_amount: "0.00",
          total_mrp: "38.57",
          profit_amount: "4.08",
          gp_amount: "10.74",
          shipping_charge: "0.00",
          created_time: "2026-04-22T15:20:00+06:00",
          redeem_point: 0,
          less_amount: "0.03",
          order_cost: "33.92",
          exchange_cost: "0.00",
          due: "38.00",
          cash_paid: "0.00",
          card_paid: "0.00",
          bkash_paid: "0.00",
          nagad_paid: "0.00",
          upay_paid: "0.00",
          city_paid: "0.00",
          nexus_paid: "0.00",
          rocket_paid: "0.00",
          ecom_cash_paid: "0.00",
          ecom_online_paid: "0.00",
          reference_invoice: null,
        },
        {
          invoice_number: "INV230426000001",
          order_source: "E-commerce",
          total_amount: "851.00",
          discount_amount: 6.6,
          vat_amount: "0.00",
          exchange_amount: "0.00",
          net_amount: "851.00",
          payment_method: "Cash On Delivery",
          paid_amount: "0.00",
          total_mrp: "858.00",
          profit_amount: "95.88",
          gp_amount: "11.27",
          shipping_charge: "0.00",
          created_time: "2026-04-23T10:26:00+06:00",
          redeem_point: 0,
          less_amount: "-0.40",
          order_cost: "755.12",
          exchange_cost: "0.00",
          due: "851.00",
          cash_paid: "0.00",
          card_paid: "0.00",
          bkash_paid: "0.00",
          nagad_paid: "0.00",
          upay_paid: "0.00",
          city_paid: "0.00",
          nexus_paid: "0.00",
          rocket_paid: "0.00",
          ecom_cash_paid: "0.00",
          ecom_online_paid: "0.00",
          reference_invoice: null,
        },
      ],
    },
    sub_total: {
      sub_total_mrp: 946.57,
      total_profit: 105.94,
      total_amount: 939,
      total_discount: 7.7,
      total_vat: 0,
      total_exchange: 0,
      total_net_amount: 939,
      paid_amount: 0,
      total_gp: 11.28,
      sub_total_cost: 833.06,
      total_exchange_cost: 0,
      total_due: 939,
      total_less: 0.13,
      total_point: 0,
      total_cash: 0,
      total_bkash_paid: 0,
      total_nagad_paid: 0,
      total_upay_paid: 0,
      total_rocket_paid: 0,
      total_city_paid: 0,
      total_nexus_paid: 0,
      total_card: 0,
      ecom_cash_paid: 0,
      ecom_online_paid: 0,
      total_shipping_charge: 0,
      both_total_info: {
        outlet_net: 0,
        outlet_due: 0,
        outlet_shipping: 0,
        ecom_net: 939,
        ecom_due: 939,
        ecom_shipping: 0,
      },
    },
  },
  status: 200,
  statusText: "",
  headers: {
    "content-length": "2783",
    "content-type": "application/json",
  },
  config: {
    transitional: {
      silentJSONParsing: true,
      forcedJSONParsing: true,
      clarifyTimeoutError: false,
    },
    adapter: ["xhr", "http"],
    transformRequest: [null],
    transformResponse: [null],
    timeout: 0,
    xsrfCookieName: "XSRF-TOKEN",
    xsrfHeaderName: "X-XSRF-TOKEN",
    maxContentLength: -1,
    maxBodyLength: -1,
    env: {},
    headers: {
      Accept: "application/json, text/plain, */*",
      Authorization:
        "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ0b2tlbl90eXBlIjoiYWNjZXNzIiwiZXhwIjoxNzc3MjQ0MzMyLCJpYXQiOjE3NzcxODQzMzIsImp0aSI6IjkxNjhjN2VjMDJkYzRkYTBiNTZkNDExN2FhMDJiMDQ4IiwidXNlcl9pZCI6IjQzMTMifQ.va5XHuRlm_y_FkN7LIqssUqM7wcGxZvbaLx9lH49ULk",
    },
    method: "get",
    url: "https://dev-backend.e-hospital.io/api/v2/report/daily-combine-sales-details-report/?from_date=2026-04-22&to_date=2026-04-26&order_type=both",
  },
  request: {},
};

export const exportSalesReport = async (data) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sales Report");

  // ===============================
  // Header Section
  // ===============================
  worksheet.mergeCells("A1:Z1");
  worksheet.getCell("A1").value = "Care-Box";
  worksheet.getCell("A1").alignment = { horizontal: "center" };

  worksheet.mergeCells("A2:Z2");
  worksheet.getCell("A2").value =
    "149/A Monipuri Para, Farmgate, Tejgaon, Dhaka-1215";
  worksheet.getCell("A2").alignment = { horizontal: "center" };

  worksheet.mergeCells("A3:Z3");
  worksheet.getCell("A3").value = "Sales Report (Combined)";
  worksheet.getCell("A3").alignment = { horizontal: "center" };

  worksheet.mergeCells("A4:Z4");
  worksheet.getCell("A4").value =
    `From: ${fromDateInLastQuery} To: ${toDateInLastQuery}`;
  worksheet.getCell("A4").alignment = { horizontal: "center" };

  // ===============================
  // Table Header (2 rows)
  // ===============================
  const headerRow1 = [
    "SL",
    "Source of Selling",
    "Invoice No.",
    "MRP",
    "Shipping Charge",
    "Disc. Amount",
    "Vat",
    "Exchange",
    "Total Cost",
    "Void Amount",
    "Net Amount",
    "Due Amount",
    "Return Cost",
    "Cash",
    "Card",
    "Bkash",
    "Nagad",
    "Rocket",
    "Upay",
    "City Pay",
    "Nexus Pay",
    "Adj. Amount",
    "Pft Amount",
    "GP%",
    "Ref. Invoice",
    "Date",
  ];

  worksheet.addRow(headerRow1);

  // Style header
  worksheet.getRow(5).font = { bold: true };

  // ===============================
  // Data Rows
  // ===============================

  const table1Data = [];
  dailyCombineSales?.data?.products_data?.results.map((prod, i) => {
    let obj = {};
    obj["SL"] = i + 1;
    obj["Source of Selling"] = prod?.order_source ? prod.order_source : "";
    obj["Invoice No."] = prod?.invoice_number ? prod.invoice_number : "";
    obj["MRP"] = isValidNumber(prod?.total_mrp) ? prod.total_mrp : "";
    obj["Shipping Charge"] = isValidNumber(prod?.shipping_charge)
      ? prod.shipping_charge
      : "";
    obj["Disc. Amount"] = isValidNumber(prod?.discount_amount)
      ? prod.discount_amount
      : "";
    obj["Vat"] = isValidNumber(prod?.vat_amount) ? prod.vat_amount : "";
    obj["Exchange"] = isValidNumber(prod?.exchange_amount)
      ? prod.exchange_amount
      : "";
    obj["Total Cost"] = isValidNumber(prod?.order_cost) ? prod.order_cost : "";
    obj["Void Amount"] = isValidNumber(prod?.exchange_cost)
      ? prod?.exchange_cost
      : "";
    obj["Net Amount"] = isValidNumber(prod?.net_amount) ? prod.net_amount : "";
    obj["Due Amount"] = isValidNumber(prod?.due) ? prod.due : "";
    obj["Return Cost"] = isValidNumber(prod?.exchange_cost)
      ? prod?.exchange_cost
      : "";
    obj["Cash"] =
      prod?.order_source && prod.order_source !== "Outlet"
        ? isValidNumber(prod?.ecom_cash_paid)
          ? prod.ecom_cash_paid
          : ""
        : isValidNumber(prod?.cash_paid)
          ? prod.cash_paid
          : "";
    obj["Card"] =
      prod?.order_source && prod.order_source !== "Outlet"
        ? isValidNumber(prod?.ecom_online_paid)
          ? prod.ecom_online_paid
          : ""
        : isValidNumber(prod?.card_paid)
          ? prod.card_paid
          : "";
    obj["Bkash"] = isValidNumber(prod?.bkash_paid) ? prod.bkash_paid : "";
    obj["Nagad"] = isValidNumber(prod?.nagad_paid) ? prod.nagad_paid : "";
    obj["Rocket"] = isValidNumber(prod?.rocket_paid) ? prod.rocket_paid : "";
    obj["Upay"] = isValidNumber(prod?.upay_paid) ? prod.upay_paid : "";
    obj["City Pay"] = isValidNumber(prod?.city_paid) ? prod.city_paid : "";
    obj["Nexus Pay"] = isValidNumber(prod?.nexus_paid) ? prod.nexus_paid : "";
    obj["Adj. Amount"] = isValidNumber(prod?.less_amount)
      ? prod.less_amount
      : "";
    obj["Pft Amount"] = isValidNumber(prod?.profit_amount)
      ? prod.profit_amount
      : "";
    obj["GP%"] = isValidNumber(prod?.gp_amount) ? prod.gp_amount : "";
    obj["Ref. Invoice"] = prod.reference_invoice ? prod.reference_invoice : "";
    obj["Date"] =
      `${convertDate(prod.created_time)} ${convertTime(prod.created_time)}`;
    table1Data.push(obj);
  });
  //   console.log("table1Data = ", table1Data);

  dailyCombineSales?.data?.products_data?.results.map((prod, i) => {
    const row = [];

    row.push(i + 1);
    row.push(prod?.order_source || "");
    row.push(prod?.invoice_number || "");
    row.push(isValidNumber(prod?.total_mrp) ? parseFloat(prod.total_mrp) : "");
    row.push(
      isValidNumber(prod?.shipping_charge)
        ? parseFloat(prod.shipping_charge)
        : "",
    );
    row.push(
      isValidNumber(prod?.discount_amount)
        ? parseFloat(prod.discount_amount)
        : "",
    );
    row.push(
      isValidNumber(prod?.vat_amount) ? parseFloat(prod.vat_amount) : "",
    );
    row.push(
      isValidNumber(prod?.exchange_amount)
        ? parseFloat(prod.exchange_amount)
        : "",
    );
    row.push(
      isValidNumber(prod?.order_cost) ? parseFloat(prod.order_cost) : "",
    );
    row.push(
      isValidNumber(prod?.exchange_cost) ? parseFloat(prod.exchange_cost) : "",
    );
    row.push(
      isValidNumber(prod?.net_amount) ? parseFloat(prod.net_amount) : "",
    );
    row.push(isValidNumber(prod?.due) ? parseFloat(prod.due) : "");
    row.push(
      isValidNumber(prod?.exchange_cost) ? parseFloat(prod.exchange_cost) : "",
    );

    // Cash
    row.push(
      prod?.order_source && prod.order_source !== "Outlet"
        ? isValidNumber(prod?.ecom_cash_paid)
          ? parseFloat(prod.ecom_cash_paid)
          : ""
        : isValidNumber(prod?.cash_paid)
          ? parseFloat(prod.cash_paid)
          : "",
    );

    // Card
    row.push(
      prod?.order_source && prod.order_source !== "Outlet"
        ? isValidNumber(prod?.ecom_online_paid)
          ? parseFloat(prod.ecom_online_paid)
          : ""
        : isValidNumber(prod?.card_paid)
          ? parseFloat(prod.card_paid)
          : "",
    );

    row.push(
      isValidNumber(prod?.bkash_paid) ? parseFloat(prod.bkash_paid) : "",
    );
    row.push(
      isValidNumber(prod?.nagad_paid) ? parseFloat(prod.nagad_paid) : "",
    );
    row.push(
      isValidNumber(prod?.rocket_paid) ? parseFloat(prod.rocket_paid) : "",
    );
    row.push(isValidNumber(prod?.upay_paid) ? parseFloat(prod.upay_paid) : "");
    row.push(isValidNumber(prod?.city_paid) ? parseFloat(prod.city_paid) : "");
    row.push(
      isValidNumber(prod?.nexus_paid) ? parseFloat(prod.nexus_paid) : "",
    );
    row.push(
      isValidNumber(prod?.less_amount) ? parseFloat(prod.less_amount) : "",
    );
    row.push(
      isValidNumber(prod?.profit_amount) ? parseFloat(prod.profit_amount) : "",
    );
    row.push(isValidNumber(prod?.gp_amount) ? parseFloat(prod.gp_amount) : "");
    row.push(prod?.reference_invoice || "");
    row.push(
      `${convertDate(prod.created_time)} ${convertTime(prod.created_time)}`,
    );

    worksheet.addRow(row);
  });

  const totalRowIndex = worksheet.lastRow.number + 1;

  worksheet.mergeCells(`A${totalRowIndex}:B${totalRowIndex}`);
  worksheet.getCell(`A${totalRowIndex}`).value = "Grand Total";

  worksheet.getRow(totalRowIndex).font = { bold: true };

  // ===============================
  // Auto Width
  // ===============================
  worksheet.columns.map((col, index) => {
    // console.log("index = ", index);
    // if (index === 0) {
    //   col.width = 7;
    // } else {
    //   col.width = 16;
    // }

    col.width = 16;
  });

  //2ed Table
  let lastIndex = worksheet.lastRow.number + 3;

  worksheet.mergeCells(`A${lastIndex}:D${lastIndex}`);
  worksheet.getCell(`A${lastIndex}`).value = "Source Wise Sales Report";
  worksheet.getCell(`A${lastIndex}`).font = { bold: true };
  worksheet.getCell(`A${lastIndex}`).alignment = {
    // horizontal: "left",
    horizontal: "center",
    vertical: "middle", // optional, looks better in merged cells
  };

  worksheet.addRow(["Name", "Amount", "Due", "Shipment"]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  // Outlet Row
  worksheet.addRow([
    "Outlet",
    isValidNumber(
      dailyCombineSales?.data?.sub_total?.both_total_info?.outlet_net,
    )
      ? dailyCombineSales.data.sub_total.both_total_info.outlet_net
      : "",

    isValidNumber(
      dailyCombineSales?.data?.sub_total?.both_total_info?.outlet_due,
    )
      ? parseFloat(-dailyCombineSales.data.sub_total.both_total_info.outlet_due)
      : "",
    isValidNumber(
      dailyCombineSales?.data?.sub_total?.both_total_info?.outlet_shipping,
    )
      ? dailyCombineSales.data.sub_total.both_total_info.outlet_shipping
      : "",
  ]);

  // E-commerce Row
  worksheet.addRow([
    "E-Commerce",
    isValidNumber(dailyCombineSales?.data?.sub_total?.both_total_info?.ecom_net)
      ? dailyCombineSales.data.sub_total.both_total_info.ecom_net
      : "",
    isValidNumber(dailyCombineSales?.data?.sub_total?.both_total_info?.ecom_due)
      ? -dailyCombineSales.data.sub_total.both_total_info.ecom_due
      : "",
    isValidNumber(
      dailyCombineSales?.data?.sub_total?.both_total_info?.ecom_shipping,
    )
      ? dailyCombineSales.data.sub_total.both_total_info.ecom_shipping
      : "",
  ]);

  // Net Amount Row
  worksheet.addRow([
    "Net Amount",
    isValidNumber(dailyCombineSales?.data?.sub_total?.both_total_info?.ecom_net)
      ? dailyCombineSales.data.sub_total.both_total_info.ecom_net
      : "",
    isValidNumber(dailyCombineSales?.data?.sub_total?.both_total_info?.ecom_due)
      ? -dailyCombineSales.data.sub_total.both_total_info.ecom_due
      : "",
    isValidNumber(
      dailyCombineSales?.data?.sub_total?.both_total_info?.ecom_shipping,
    )
      ? dailyCombineSales.data.sub_total.both_total_info.ecom_shipping
      : "",
  ]);

  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  //3ed Table

  lastIndex = worksheet.lastRow.number + 3;
  worksheet.getCell(`A${lastIndex}`).value = "Cash Amount";
  worksheet.getCell(`B${lastIndex}`).value = isValidNumber(
    dailyCombineSales?.data?.sub_total?.total_cash,
  )
    ? dailyCombineSales.data.sub_total.total_cash
    : "";
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Card Amount",
    isValidNumber(dailyCombineSales?.data?.sub_total?.total_card)
      ? dailyCombineSales.data.sub_total.total_card
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Bkash",
    isValidNumber(dailyCombineSales?.data?.sub_total?.total_bkash_paid)
      ? dailyCombineSales.data.sub_total.total_bkash_paid
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Nagad",
    isValidNumber(dailyCombineSales?.data?.sub_total?.total_nagad_paid)
      ? dailyCombineSales.data.sub_total.total_nagad_paid
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Rocket",
    isValidNumber(dailyCombineSales?.data?.sub_total?.total_rocket_paid)
      ? dailyCombineSales.data.sub_total.total_rocket_paid
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Upay",
    isValidNumber(dailyCombineSales?.data?.sub_total?.total_upay_paid)
      ? dailyCombineSales.data.sub_total.total_upay_paid
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "City Pay",
    isValidNumber(dailyCombineSales?.data?.sub_total?.total_city_paid)
      ? dailyCombineSales.data.sub_total.total_city_paid
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Nexus Pay",
    isValidNumber(dailyCombineSales?.data?.sub_total?.total_nexus_paid)
      ? dailyCombineSales.data.sub_total.total_nexus_paid
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Net Amount",
    dailyCombineSales?.data?.sub_total?.total_net_amount,
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Due Amount",
    -dailyCombineSales?.data?.sub_total?.total_due,
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Return Amount",
    -dailyCombineSales?.data?.sub_total?.total_exchange_cost,
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  worksheet.addRow([
    "Grand Total",
    dailyCombineSales?.data
      ? roundTo(
          parseFloat(dailyCombineSales?.data?.sub_total?.total_net_amount) -
            ((parseFloat(dailyCombineSales?.data?.sub_total?.total_due)
              ? parseFloat(dailyCombineSales?.data?.sub_total?.total_due)
              : 0) +
              (parseFloat(
                dailyCombineSales?.data?.sub_total?.total_exchange_cost,
              )
                ? parseFloat(
                    dailyCombineSales?.data?.sub_total?.total_exchange_cost,
                  )
                : 0)),
        )
      : "",
  ]);
  worksheet.getRow(worksheet.lastRow.number).font = { bold: true };

  // ===============================
  // Download
  // ===============================
  const buffer = await workbook.xlsx.writeBuffer();
  saveAs(new Blob([buffer]), "Sales-Report.xlsx");
};
