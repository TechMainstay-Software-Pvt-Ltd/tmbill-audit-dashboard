"use client";

import { useDeferredValue, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

type Invoice = {
  id:string; billNo:string; date:string; time:string; items:number; subtotal:number; discount:number; vat:number;
  charges:number; total:number; status:string; orderType:string; paymentMode:string; user:string; source:string;
  discountReason:string; itemLineDiscount:number; discountVariance:number; lineGross:number; lineNetExport:number; hasAnomaly:boolean; table:string; guests:number;
  taxableGrossInclVat:number; taxControl:number; taxVariance:number; taxableNetExVat:number; taxAssessment:string;
  canonicalItemDiscount:number;canonicalOrderDiscount:number;canonicalGrossBeforeDiscount:number;canonicalTotalDiscount:number;
  canonicalTaxableInclVat:number;canonicalVat:number;canonicalNetExVat:number;
  discountAssessment:string;vatAssessment:string;controlStatus:string;
};
type Item = {
  rowId:number; invoiceId:string; billNo:string; date:string; itemId:string; name:string; category:string; orderType:string;
  qty:number; actualPrice:number; expectedCurrentPrice:number|null; priceVariance:number|null; priceStatus:string; vatRate:number;
  gross:number; lineDiscount:number; netAmount:number; vatBase:number; invoiceDiscount:number; invoiceTotal:number; anomaly:boolean;
  canonicalItemDiscount:number;canonicalOriginalGross:number;canonicalOrderDiscount:number;canonicalNetInclVat:number;canonicalVat:number;canonicalNetExVat:number;
  monthlyAvgGrossUnit:number;monthlyAvgSoldInclVat:number;monthlyAvgSoldExVat:number;orderTypeAvgGrossUnit:number;orderTypeAvgSoldInclVat:number;orderTypeAvgSoldExVat:number;effectiveDiscountRate:number;
};
type Category = {name:string;qty:number;gross:number;net:number;discount:number;itemDiscount:number;orderDiscount:number;vat:number;netExVat:number;dashboardMix:number};
type AuditData = {
  summary:Record<string,number|string>; invoices:Invoice[]; items:Item[]; categories:Category[];
  topItems:Array<{name:string;qty:number;gross:number;net:number;discount:number}>;
};
type PaymentReconciliation={summary:Record<string,any>;exceptions:any[]};
type SeparateAudit={summary:Record<string,any>;invoices:any[];orderTypes:any[];payments:any[];sources:any[];orderPayment:any[]};
type KotAudit={
  summary:Record<string,any>;bills:any[];events:any[];cancellations:any[];edited:any[];
  complimentary:any[];topCancelledItems:any[];reportAudit:any[];
};
type BusinessInsights={
  summary:Record<string,number>;turnaroundByType:any[];invoiceTurnaround:Record<string,any>;
  turnaroundRows:any[];discountSummary:any[];discountDetails:any[];hourly:any[];daily:any[];
  weekday:any[];weekly:any[];operatingHighlights:Record<string,any>;
};

function auditUploadedSalesRows(rows:any[][],fileName:string):SeparateAudit {
  const headers=(rows[0]||[]).map(x=>String(x??"").trim());
  const at=(name:string)=>headers.indexOf(name);
  const n=(r:any[],name:string)=>Number(r[at(name)]||0);
  const t=(r:any[],name:string)=>String(r[at(name)]??"");
  const deliveryCols=headers.map((h,i)=>h==="Delivery Charges"?i:-1).filter(i=>i>=0);
  const dStart=at("specialdisc5%"),dEnd=at("Complementary");
  const raw=rows.slice(1).filter(r=>String(r[at("Order Id")]||"").startsWith("10"));
  if(!raw.length||at("Total")<0||at("Total Tax")<0)throw new Error("This is not a compatible TMBill Sales Report. Required headers: Order Id, Bill No, Sub Total, Total Discount, Total Tax, Total and Status.");
  const invoiceRows=raw.map(r=>{
    const subtotal=n(r,"Sub Total"),discount=n(r,"Total Discount"),vat=n(r,"Total Tax"),charges=n(r,"Total Charges"),total=n(r,"Total"),rounded=n(r,"Rounded Amount");
    const discountComponents=dStart>=0&&dEnd>=dStart?r.slice(dStart,dEnd+1).reduce((a:number,v:any)=>a+Number(v||0),0):discount;
    const taxable=subtotal-discount,vatControl=Math.round((taxable*5/105+Number.EPSILON)*100)/100;
    const chargeParts=(deliveryCols.length?Number(r[deliveryCols[0]]||0):0)+n(r,"Item Level Charges");
    const issues:string[]=[];
    if(Math.abs(discountComponents-discount)>.01)issues.push("Discount components do not equal Total Discount");
    if(Math.abs(total-(subtotal-discount+charges+rounded))>.01)issues.push("Invoice arithmetic mismatch");
    if(Math.abs(vat-vatControl)>.02)issues.push("VAT differs from aggregate 5/105 control");
    if(Math.abs(chargeParts-charges)>.01)issues.push("Charge components do not equal Total Charges");
    if(charges&&n(r,"Tax On Charges")===0)issues.push("Charges present with zero Tax On Charges");
    if(t(r,"Payment Mode").includes(",")||t(r,"Payment Mode").toLowerCase()==="split")issues.push("Split payment amounts unavailable");
    return {id:t(r,"Order Id"),billNo:String(r[at("Bill No")]??""),date:t(r,"Order Date"),time:t(r,"Order Time"),items:n(r,"No. Of Items"),subtotal,discount,discountComponents,discountVariance:discountComponents-discount,vat,vatControl,vatVariance:vat-vatControl,charges,chargeParts,taxOnCharges:n(r,"Tax On Charges"),total,totalControl:subtotal-discount+charges+rounded,totalVariance:total-(subtotal-discount+charges+rounded),status:t(r,"Status"),user:t(r,"User"),orderType:t(r,"Type"),paymentMode:t(r,"Payment Mode"),source:t(r,"Order Source"),guests:n(r,"Cover/Guest"),table:t(r,"Table Name"),discountReason:t(r,"Discount Reason"),issues,assessment:issues.length?"review":"pass"};
  });
  const fulfilled=invoiceRows.filter(x=>x.status.toLowerCase()==="fulfilled");
  const aggregate=(key:string)=>{const m=new Map<string,any>();for(const x of fulfilled){const name=x[key]||"Unmapped",z=m.get(name)||{name,bills:0,subtotal:0,discount:0,vat:0,charges:0,total:0};z.bills++;for(const k of ["subtotal","discount","vat","charges","total"])z[k]+=x[k];m.set(name,z)}return [...m.values()].map(x=>({...x,subtotal:+x.subtotal.toFixed(2),discount:+x.discount.toFixed(2),vat:+x.vat.toFixed(2),charges:+x.charges.toFixed(2),total:+x.total.toFixed(2)}))};
  const cm=new Map<string,any>();for(const x of fulfilled){const key=`${x.orderType}|||${x.paymentMode}`,z=cm.get(key)||{orderType:x.orderType,paymentMode:x.paymentMode,bills:0,total:0,vat:0};z.bills++;z.total+=x.total;z.vat+=x.vat;cm.set(key,z)}
  const orderPayment=[...cm.values()].map(x=>{const agg=["Talabat","Keeta","Deliveroo","Careem","Noonfood","Smiles"].includes(x.orderType),expected=agg?x.paymentMode.toLowerCase()===x.orderType.toLowerCase():["Card","Cash","Card,Cash"].includes(x.paymentMode);return {...x,total:+x.total.toFixed(2),vat:+x.vat.toFixed(2),assessment:expected?"pass":"review"}});
  const sum=(k:string)=>+fulfilled.reduce((a,x)=>a+Number(x[k]||0),0).toFixed(2);
  const bills=invoiceRows.map(x=>Number(x.billNo)).filter(Number.isFinite);
  const dates=invoiceRows.map(x=>x.date).filter(Boolean);
  const summary:any={file:fileName,periodFrom:dates[0]||"—",periodTo:dates[dates.length-1]||"—",rows:invoiceRows.length,fulfilled:fulfilled.length,statusCounts:{},billFrom:Math.min(...bills),billTo:Math.max(...bills),subtotal:sum("subtotal"),discount:sum("discount"),vat:sum("vat"),charges:sum("charges"),taxOnCharges:sum("taxOnCharges"),total:sum("total"),netExVat:+(sum("total")-sum("vat")).toFixed(2),vatControl:sum("vatControl"),vatControlVariance:+(sum("vat")-sum("vatControl")).toFixed(2),totalControlVariance:sum("totalVariance"),chargeComponentGap:+fulfilled.reduce((a,x)=>a+x.chargeParts-x.charges,0).toFixed(2),reviewBills:fulfilled.filter(x=>x.assessment==="review").length,splitUnverifiable:fulfilled.filter(x=>x.issues.includes("Split payment amounts unavailable")).length,chargeBills:fulfilled.filter(x=>x.charges).length,channelMappingReviews:orderPayment.filter(x=>x.assessment==="review").length};
  return {summary,invoices:invoiceRows,orderTypes:aggregate("orderType"),payments:aggregate("paymentMode"),sources:aggregate("source"),orderPayment};
}

const tabs = [
  ["overview","Control room"],["bills","Bill-to-bill"],["items","Items & prices"],["discounts","Discounts"],
  ["operations","Item & KOT audit"],["payments","Payments"],["crosscheck","Report controls"],["separate","Audit any sales report"],["tax","VAT audit"],["categories","Categories"],["zreport","Today / Z report"],["reports","Ideal reports"],["menu","Menu engineering"],["guide","How to read"]
];
const money = (v:number) => `\u20C3 ${new Intl.NumberFormat("en-AE",{minimumFractionDigits:2,maximumFractionDigits:2}).format(v)}`;
const printMoney = (v:number) => `AED ${new Intl.NumberFormat("en-AE",{minimumFractionDigits:2,maximumFractionDigits:2}).format(v)}`;
const num = (v:number) => new Intl.NumberFormat("en-AE").format(v);
const isZeroDisplay = (v:any) => typeof v==="number"?Math.abs(v)<.000001:["0","0.00",money(0),printMoney(0)].includes(String(v??"").trim());
const tabIcons:Record<string,string>={overview:"⌂",bills:"▤",items:"◇",discounts:"%",operations:"⌁",payments:"◉",crosscheck:"⇄",separate:"↑",tax:"VAT",categories:"▦",zreport:"Z",reports:"▧",menu:"★",guide:"?"};

const reportSpecs = [
 ["Sales / Z Summary","Branch × business date × shift","The master financial control using the standardized sales bridge from Gross Item Sales to Actual Sales incl. VAT."],
 ["Cash Audit","Drawer × cashier × shift","Opening float through counted cash, with every cash movement and variance."],
 ["Order Type Summary","Period × order type","Dine-in, pickup, delivery and aggregator mix using the same canonical invoice sales."],
 ["Payment Type Summary","Period × tender","One row per tender allocation, including every part of split payments."],
 ["Discount Summary","Rule × reason × user","Discounts and Complimentary shown separately, with eligibility, reason, user and approval exceptions."],
 ["Expense Summary","Expense transaction","Operational expense, input VAT, payment method, approval and reference."],
 ["Bill Summary","One row per invoice","The single invoice ledger from which all financial summaries are produced."],
 ["Delivery Boy Summary","Period × delivery employee","Orders, sales, delivery time, cash collections and handover variance."],
 ["Waiter Summary","Period × waiter","Bills, guests, average check, voids, tips and table turns."],
 ["Product Group Summary","Period × product group","Gross Item Sales, Discounts, Complimentary, Net Item Sales incl./excl. VAT and mix."],
 ["Kitchen Department Summary","Period × kitchen","Production quantity, value, cancellations, preparation time and late tickets."],
 ["Category Summary","Period × category","The standardized sales bridge by category; totals plus Charges incl. VAT must equal Actual Sales incl. VAT."],
 ["Sold Items Summary","Item × order type","The standardized sales bridge by item and order type, with realized price and menu-price variance."],
 ["Cancel Items Summary","One status event","Invoice/item/KOT identity, original value, actor, approver, reason and time."],
 ["Wallet Summary","One wallet event","Credits, debits, expiry, invoice reference and liability balance."],
 ["Due Payment Received","One receipt allocation","Receipt against original invoice with tender, reference and remaining balance."],
 ["Due Payment Receivable","One open invoice","Customer balance and aging from supplied but unpaid invoices."],
 ["Payment Variance","Period × tender/source","POS versus processor, bank and aggregator settlement reconciliation."],
 ["Currency Denominations","Drawer × denomination","Counted cash by denomination and verifier."],
 ["Order Source Summary","Period × source","Channel economics, cancellations, discounts, check average and settlement variance."]
];
const reportHeaders:Record<string,string[]> = {
 "Sales / Z Summary":["Business date","Shift","Bills","Gross Item Sales","Discounts","Complimentary","Net Item Sales incl. VAT","VAT","Net Item Sales excl. VAT","Charges incl. VAT","Actual Sales incl. VAT","Paid","Due","Variance"],
 "Cash Audit":["Shift","Drawer","Cashier","Opening float","Cash sales","Due received","Expenses","Expected","Counted","Variance"],
 "Order Type Summary":["Order type","Bills","Guests","Gross Item Sales","Discounts","Complimentary","Net Item Sales incl. VAT","VAT","Net Item Sales excl. VAT","Charges incl. VAT","Actual Sales incl. VAT","Average check","Mix %"],
 "Payment Type Summary":["Tender","Allocations","Invoices","Collected","Refunds","Net collected","Mix %","Variance"],
 "Discount Summary":["Title","Classification","Reason","User","Bills","Eligible Gross","Amount","Effective %","Exceptions"],
 "Expense Summary":["Date","Expense ID","Category","Supplier","Net","Input VAT","Gross","Tender","Approver","Status"],
 "Bill Summary":["Bill","Order ID","Supply time","Status","Order type","Gross Item Sales","Discounts","Complimentary","Net Item Sales incl. VAT","VAT","Net Item Sales excl. VAT","Charges incl. VAT","Actual Sales incl. VAT","Paid","Due"],
 "Delivery Boy Summary":["Driver","Orders","Delivered","Cancelled","Actual Sales incl. VAT","Cash collected","Due","Avg minutes","Late %","Variance"],
 "Waiter Summary":["Waiter","Bills","Guests","Actual Sales incl. VAT","Discounts","Average check","Voids","Tips","Table turns"],
 "Product Group Summary":["Product group","Qty","Gross Item Sales","Discounts","Complimentary","Net Item Sales incl. VAT","VAT","Net Item Sales excl. VAT","Mix %"],
 "Kitchen Department Summary":["Kitchen","Tickets","Lines","Qty","Net Item Sales incl. VAT","Cancelled qty","Cancel value","Avg prep","Late %"],
 "Category Summary":["Category","Qty","Gross Item Sales","Discounts","Complimentary","Net Item Sales incl. VAT","VAT","Net Item Sales excl. VAT","Mix %"],
 "Sold Items Summary":["SKU","Item","Order type","Qty","Menu price","Actual price","Gross Item Sales","Discounts","Complimentary","Net Item Sales incl. VAT","VAT","Net Item Sales excl. VAT","Price variance"],
 "Cancel Items Summary":["Event","Time","Invoice","KOT","Item","Qty","Original value","From","To","Reason","User","Approver"],
 "Wallet Summary":["Event","Time","Wallet","Type","Invoice","Credit","Debit","Expiry","Balance","Status"],
 "Due Payment Received":["Receipt","Received at","Invoice","Customer","Opening due","Received","Tender","Reference","Remaining"],
 "Due Payment Receivable":["Invoice","Supply date","Customer","Gross","Paid","Receipts","Credits","Open due","Age","Bucket"],
 "Payment Variance":["Date","Tender/source","POS net","Processor net","Fees","Expected deposit","Actual","Timing","Variance","Owner"],
 "Currency Denominations":["Date","Shift","Drawer","Currency","Denomination","Quantity","Amount","Counter","Verifier"],
 "Order Source Summary":["Source","Orders","Fulfilled","Cancelled","Discounts","Complimentary","Charges incl. VAT","VAT","Actual Sales incl. VAT","Average check","Cancel %","Mix %"]
};
const paymentMatrix=[
 {orderType:"Dine In",orders:684,card:29428.20,cash:5702.50,talabat:117,deliveroo:0,keeta:0,charges:50,vat:1676.78,sales:35247.70,exported:34371.80},
 {orderType:"Pickup",orders:488,card:16713.45,cash:3549.40,talabat:0,deliveroo:0,keeta:0,charges:0,vat:965.29,sales:20262.85,exported:19956.85},
 {orderType:"Delivery",orders:52,card:7179.50,cash:1478,talabat:0,deliveroo:0,keeta:0,charges:531,vat:387,sales:8657.50,exported:8657.50},
 {orderType:"Talabat",orders:577,card:0,cash:0,talabat:42484.10,deliveroo:0,keeta:65.60,charges:0,vat:2026.54,sales:42549.70,exported:42549.70},
 {orderType:"Keeta",orders:15,card:0,cash:0,talabat:0,deliveroo:0,keeta:1455,charges:0,vat:69.31,sales:1455,exported:1455}
];
const sourceAudits=[
 {file:"daily-sales-report…xlsx",name:"Daily Sales Report",status:"fail",issue:"Export is effectively blank: only title, Discount, Total and an empty TOTAL row.",fix:"Populate governed Z-report sections from the invoice ledger; fail export generation when detail sections are empty.",headers:"DAY, DATE, BRANCH, DISCOUNT, TOTAL"},
 {file:"day-wise-consolidated-report…xlsx",name:"Day-wise Consolidated",status:"warn",issue:"Daily rows reconcile, but the workbook includes a grand-total row; naïve column sums double every value. Total Sale and Net Total are duplicate concepts.",fix:"Tag row_type, keep totals outside the data table, and define Net Sales as excluding VAT while Total Sale includes VAT and charges.",headers:"Order Date, Store Name, Total Sale, Sub Total, Discounts, Refunds, Refunded Tax, Net Sales, Taxes, Charges, Net Total, Tips, Total Orders, Average Per Order, Total Customers, Return Quantity, Round Off"},
 {file:"dsr-bill-no-of-series…xlsx",name:"Bill Number Series",status:"fail",issue:"All receipt-series fields are blank or zero despite fulfilled bill numbers #2313–#4141 and 1,816 bills.",fix:"Populate series start/end, issued count, void count, actual count and explicitly list sequence gaps.",headers:"State, Store Name, Starting No., Ending No., Count of Receipt No., Count of Void Invoices, Count of Actual Receipt"},
 {file:"dsr-bill-wise-report…xlsx",name:"DSR Bill-wise",status:"warn",issue:"Core sales and VAT reconcile, but Sale-5% is actually net sales excluding VAT; Net Sale repeats it. Delivery Charges show ⃃ 520 while invoice Total Charges show ⃃ 581.",fix:"Rename tax bases, add item/order discount separately, show charge type and tax, and retain one row_type-free invoice table.",headers:"Date, Time, Order Id, Ticket, Nature of Supply, Sale-5%, 5% VAT, Net Sale, Total Sale, Delivery Charges, Total Charges, Discount, Status, card, cash, deliveroo, keeta, talabat, Payment Mode, Guest"},
 {file:"dsr-day-wise…xlsx",name:"DSR Day-wise",status:"warn",issue:"Payment and sales totals reconcile, but Sale-5% and Net Sale repeat the same ⃃ 103,047.83 and do not explain gross-before-discount or charge VAT.",fix:"Show original gross, item discount, order discount, taxable ex VAT, VAT, taxable incl VAT, charges and revenue as separate columns.",headers:"Date, Ticket count, card, cash, deliveroo, keeta, talabat, Sale-5%, VAT-5%, Net Sale, Delivery Charges, Round off, Total Sale"},
 {file:"dsr-item-wise…xlsx",name:"DSR Item-wise",status:"fail",issue:"Line Discount mixes allocated order discounts with true item discounts. #2990 exports Fattoush at price 0, discount 13 and negative net/tax base. Line net totals do not reconcile to invoice sales.",fix:"Add original unit price, item discount, allocated order discount, taxable incl/ex VAT and line VAT; prohibit negative taxable lines.",headers:"Date, Receipt IDs, Category, Item ID, Description, Sales Type, VAT %, Quantity, Price, Delivery Charges, Net Amount, Line Discount, Sale-Exempt, VAT Base Amount, Tax Product Group"},
 {file:"dsr-month-wise-report…xlsx",name:"DSR Month-wise",status:"warn",issue:"Core invoice totals reconcile, but Delivery/Total Charges are ⃃ 520 instead of the invoice-ledger ⃃ 581, leaving ⃃ 61 unclassified.",fix:"Use the charge ledger, one column per charge type plus total charge VAT; rename Sale-5% to taxable sales excluding VAT.",headers:"Period, Tickets, Sale-5%, 5% VAT, Net Sale, Total Sale, Delivery Charges, Total Charges, Discount, payment columns"},
 {file:"month-wise-sales…xlsx",name:"Order Type Month-wise",status:"pass",issue:"Order-type subtotal, discount, VAT and total reconcile to the invoice ledger. It lacks charges, item discounts and validation columns.",fix:"Retain as a summary but add item/order discount split, charges, net ex VAT, VAT variance and reconciliation status.",headers:"Description, Sum of Sub Total, Sum of Total Discount, Sum of Actual VAT 5%, Sum of Total"},
 {file:"Morbido Express Restaurant Menu.xlsx",name:"Menu Master",status:"warn",issue:"Useful current menu snapshot, but no effective-from/to versioning and no immutable cross-report item key in every transaction export.",fix:"Version prices and tax groups by item ID, order type and effective dates; add recipe/food cost for menu engineering.",headers:"Short Code, Title, Tax Product Group, Category, Kitchen Dept, Base Item Price, status and platform fields"},
 {file:"Morbido Express Restaurant Multi Price File Format.xlsx",name:"Multi-price Master",status:"warn",issue:"Order-type prices exist, but only as a current snapshot. Historical June mismatches cannot be proven wrong from this file.",fix:"Store immutable item ID, order type, price, tax inclusion and effective-from/to dates.",headers:"Short Code, Item Name, Base Item Price, Quick Bill, Dine In, Pickup, Delivery, Talabat, Deliveroo, Keeta"},
 {file:"order-type-day-wise-report…xlsx",name:"Order Type Day-wise",status:"fail",issue:"Delivery rows violate Subtotal − Discount = Total because charges are embedded inconsistently; period grand subtotal/discount/total do not match the invoice ledger.",fix:"Keep item sales and charges separate, include VAT, and reconcile every order-type/day row to invoice IDs.",headers:"Date, Day, per-order-type Sub Total, Total Discount, Total, Grand Totals"},
 {file:"sales-report…xlsx",name:"Sales Report",status:"warn",issue:"Best invoice ledger and core control source, but it contains 12 identically named Delivery Charges headers, hiding charge identity. Tax On Charges is zero while Total Charges are ⃃ 581.",fix:"Unique charge_type columns or a child charge table; show item/order discounts separately and validate charge tax treatment.",headers:"Order ID, Bill, dates, items, Sub Total, discount rules, Total Discount, VAT, tax modes, Total Tax, repeated charges, Item Level Charges, Total Charges, Tax On Charges, Total, Status, Type, Payment, Source, Guests"},
 {file:"sales-report-with-items…xlsx",name:"Sales Report With Items",status:"fail",issue:"Hierarchical invoice and item rows share columns and include a summary row, making flat sums duplicate invoice totals and obscuring discount allocation.",fix:"Export separate Invoice and Invoice Line sheets joined by Order ID; lines must carry both discount levels and VAT.",headers:"Invoice financial headers followed by HSN, Item Name, Price, Qty, line Total, Product Group, Category, Return Item"},
 {file:"sales-report-with-product-group-details…xlsx",name:"Product Group Details",status:"fail",issue:"Multiple columns are all named Delivery Charges; summary/footer rows sit inside the data range; Taxable Amount and VAT summary labels are inconsistent.",fix:"One normalized product-group table with unique headers, row_type, corrected allocated discounts, line VAT and explicit charge allocation policy.",headers:"Bill Date, Order ID, Bill, Qty, Subtotal, Taxable Amount, Discount, VAT 5%, repeated Delivery Charges, Round Off, Total Amount, 0%, Status"},
 {file:"simplified-day-wise-dsr…xlsx",name:"Simplified Day-wise DSR",status:"warn",issue:"Daily and period totals broadly reconcile, but Total Taxable Amount is gross pre-discount, not the taxable amount used for VAT. Credit combines aggregator tenders.",fix:"Rename gross pre-discount, show taxable base after discounts, split aggregator tenders and show charges/tax separately.",headers:"Day, Date, Invoice count, Total Taxable Amount, Non Taxable, Total Tax, tender sales/tax/ex-tax, Discount, Charges, Grand Total"},
 {file:"tax-submission-payment-report…xlsx",name:"Tax Submission Payment-wise",status:"fail",issue:"Total Collection is ⃃ 106,990.85, short of invoice/payment sales by ⃃ 1,181.90. Duplicate VAT headers make the grain ambiguous.",fix:"Build from payment-allocation facts, not a primary payment label; one row per invoice × tender with allocated VAT only for analytics, never as the filing basis.",headers:"Order Type, Order Count, repeated VAT columns, Card/Cash/Talabat/Deliveroo/Keeta/Due collections, Total Collection"},
 {file:"tax-submission-report…xlsx",name:"Tax Submission Bill-wise",status:"fail",issue:"Tender sales total only ⃃ 106,990.85 while Total Sales is ⃃ 108,172.75. Payment columns omit secondary split allocations although invoice VAT total is correct.",fix:"Separate supply/VAT filing ledger from payment reconciliation. Explode every tender allocation and validate tender sum = invoice total.",headers:"Bill Number, Order Type, per-tender VAT and Sales, Total VAT, Total Sales"},
 {file:"day-wise-summary-report…xlsx",name:"New Day-wise Summary",status:"pass",issue:"This is the strongest day-level control received: it reconciles 1,816 bills, ⃃ 118,053 item subtotal, ⃃ 10,461.25 header discounts, ⃃ 5,124.92 VAT, ⃃ 581 charges and ⃃ 108,172.75 revenue. It omits at least the confirmed ⃃ 13 item adjustment on #2990; completeness cannot be certified from this summary.",fix:"Retain its daily bridge but add original transaction gross, separately stored item discounts, order discounts and a completeness control sourced from immutable invoice lines.",headers:"Outlet, Date, Bill Range, Total Bill, Total Discount, Total Tax, Total Charges, Item Total, Net Sales, Grand Total"},
 {file:"discount-report…xlsx",name:"New Discount Report",status:"warn",issue:"Order-type discount total ⃃ 10,461.25 reconciles to invoice header discounts, but the report has no discount rule, bill, item/order level, approver or eligibility fields. It omits the confirmed ⃃ 13 item adjustment and cannot reveal partial item discounts.",fix:"Add invoice and item grain, original transaction-time price, discount level, eligible base, rule, reason, user and approval. Treat ⃃ 10,474.25 as a confirmed minimum until line-level completeness is available.",headers:"Order From, Discount Amount, Status"},
 {file:"item-wise-report…xlsx",name:"New Item-wise Report",status:"fail",issue:"All 3,367 item rows report Discount = zero. Bill #2990 shows Fattoush price zero and discount zero, so the confirmed ⃃ 13 item adjustment and all allocated order discounts disappear. Other partial item discounts cannot be distinguished from price-version changes.",fix:"Export original transaction-time price, item discount, allocated order discount, sold incl/ex VAT, line VAT and price-version ID. Join by immutable Order ID and item-line ID.",headers:"Order ID, Title, Quantity, Price, Discount Name, Discount, Discount Reason, Item Note, Bill Number, Order Date, Table Name"},
 {file:"order-state-transition-report…xlsx",name:"Order State Transition",status:"warn",issue:"Only 1,229 offline-type records are present (Dine In/Pickup/Delivery); all aggregator orders are absent. Food Ready and Dispatched timestamps are entirely empty, and some rows contain both completed and canceled timestamps.",fix:"Include every order source, use one immutable status-event row per transition, and validate mutually exclusive terminal states.",headers:"Order ID, Platform, Order Type, Store, Placed, Acknowledged, Food Ready, Dispatched, Completed, Canceled, Username, Source, Duration"},
 {file:"shift-wise-report…xlsx",name:"Shift-wise Report",status:"fail",issue:"Report header dates incorrectly show 13 Feb 2004 to 03 Apr 2006. Seven June shifts have zero Total Sale and the shift total is only ⃃ 85,346.45 rather than ⃃ 108,172.75.",fix:"Repair date formatting/source parameters; allocate every invoice to a shift by settlement timestamp and require shift totals to reconcile to day sales.",headers:"Shift Start, Start User, Shift End, End User, Closing Balance, Opening Balance, Expense, Total Sale, Current Closing Balance, Comments"},
 {file:"start-close-day-report…xlsx",name:"Start / Close Day Report",status:"fail",issue:"June sales total reconciles, but invoice ranges do not align with the actual June bill series and 16 days have non-zero cash differences totaling AED -10,189.20, often because Close Day Amount is zero.",fix:"Separate expected cash from counted cash, require count/approval before close, reconcile invoice range to issued bills and show explained versus unexplained variance.",headers:"Start Day, End Day, users, Opening/Closing Balance, Expense, Total Sale, Invoice Range, Comments, Close Day Amount, Cash Difference"},
];

function Badge({tone="neutral",children}:{tone?:string;children:React.ReactNode}) {
  return <span className={`badge ${tone}`}>{children}</span>
}
function ChannelBadge({name}:{name:string}) {
  const key=channelKey(name);
  return <span className={`channelBadge channel-${key}`}>{name||"Unclassified"}</span>
}
const channelNames=["Dine In","Pickup","Delivery","Talabat","Noon Food","Noonfood","Deliveroo","Careem","Smiles","Keeta","Card","Cash","Card,Cash","Split","Wallet","Due","Unclassified"];
const channelHex:Record<string,string>={"Dine In":"#3668A9",Pickup:"#7B57A5",Delivery:"#2F9470",Talabat:"#F58220",TALABAT:"#F58220","Noon Food":"#FFDB00",Noonfood:"#FFDB00",Deliveroo:"#00CDBC",Careem:"#36A853",Smiles:"#7C3F98",Keeta:"#FFD800",Card:"#386C9B",Cash:"#3F8A5D","Card,Cash":"#52677D",Split:"#52677D",Wallet:"#A36B12",Due:"#B43B35",Unclassified:"#68756F"};
function channelKey(name:any) {
  return String(name||"unclassified").trim().toLowerCase().replace(/[,/]+/g,"-").replace(/\s+/g,"-");
}
function isChannelValue(value:any) {
  const normalized=String(value??"").trim().toLowerCase().replace(/\s+/g," ");
  return channelNames.some(x=>x.toLowerCase()===normalized);
}
function isChannelColumn(c:GridCol) {
  return c.channel===true||/(order type|payment|tender|collection mode|order source)/i.test(`${c.key} ${c.label}`);
}
function Kpi({label,value,sub,tone="good",onClick}:{label:string;value:string;sub:string;tone?:string;onClick?:()=>void}) {
  return <article className={`kpi ${tone} ${onClick?"clickableKpi":""}`} onClick={onClick} onKeyDown={e=>{if(onClick&&(e.key==="Enter"||e.key===" ")){e.preventDefault();onClick()}}} role={onClick?"button":undefined} tabIndex={onClick?0:undefined}><div className="kpiLabel">{label}</div><div className="kpiValue">{value}</div><div className="kpiSub">{sub}</div>{onClick&&<span className="kpiAction">View detail →</span>}</article>
}
function SalesReconciliation({grossItems,totalReductions,complimentary,vat,charges,actualSales}:{grossItems:number;totalReductions:number;complimentary:number;vat:number;charges:number;actualSales:number}) {
  const discounts=+(totalReductions-complimentary).toFixed(2);
  const netItemsInclVat=+(grossItems-discounts-complimentary).toFixed(2);
  const netItemsExVat=+(netItemsInclVat-vat).toFixed(2);
  const reconciled=+(netItemsInclVat+charges).toFixed(2);
  const variance=+(reconciled-actualSales).toFixed(2);
  return <div className="reportReconciliation" aria-label="Report to actual sales reconciliation">
    <div><span>Gross Item Sales</span><b>{money(grossItems)}</b><small>Before reductions</small></div><i>−</i>
    <div><span>Discounts</span><b>{money(discounts)}</b><small>Order and promotional reductions</small></div><i>−</i>
    <div><span>Complimentary</span><b>{money(complimentary)}</b><small>Authorized free items</small></div><i>=</i>
    <div><span>Net Item Sales incl. VAT</span><b>{money(netItemsInclVat)}</b></div>
    <div><span>VAT included</span><b>{money(vat)}</b><small>Disclosure only; do not deduct again from Actual Sales</small></div>
    <div><span>Net Item Sales excl. VAT</span><b>{money(netItemsExVat)}</b></div><i>+</i>
    <div><span>Charges incl. VAT</span><b>{money(charges)}</b></div><i>=</i>
    <div className="reconciledTotal"><span>Actual Sales incl. VAT</span><b>{money(reconciled)}</b></div>
    <div className={Math.abs(variance)<=.01?"reconPass":"reconFail"}><span>Variance to invoice sales</span><b>{money(variance)}</b><small>{Math.abs(variance)<=.01?"Matched":"Review required"}</small></div>
  </div>
}
function Info({title,children,tone="blue"}:{title:string;children:React.ReactNode;tone?:string}) {
  return <div className={`info ${tone}`}><strong>{title}</strong><div>{children}</div></div>
}
function Bar({value,max,color="#17a673"}:{value:number;max:number;color?:string}) {
  return <div className="bar"><i style={{width:`${Math.max(1,value/max*100)}%`,background:color}} /></div>
}
function minutesLabel(value:number|null|undefined) {
  if(value===null||value===undefined)return "Not captured";
  const hours=Math.floor(value/60),minutes=Math.round(value%60);
  return hours?`${hours}h ${minutes}m`:`${minutes} min`;
}
function MetricLine({rows,valueKey,labelKey,color="#187554"}:{rows:any[];valueKey:string;labelKey:string;color?:string}) {
  const max=Math.max(1,...rows.map(x=>Number(x[valueKey]||0)));
  const points=rows.map((x,i)=>`${rows.length===1?50:i/(rows.length-1)*100},${38-Number(x[valueKey]||0)/max*34}`).join(" ");
  const area=`0,38 ${points} 100,38`;
  return <div className="metricLine"><svg viewBox="0 0 100 42" preserveAspectRatio="none" role="img"><path d="M0 38H100" className="axis"/><polygon points={area} fill={color} opacity=".09"/><polyline points={points} fill="none" stroke={color} strokeWidth="1.7" vectorEffect="non-scaling-stroke"/>{rows.map((x,i)=><circle className="metricPoint" key={i} cx={rows.length===1?50:i/(rows.length-1)*100} cy={38-Number(x[valueKey]||0)/max*34} r="1.5" fill={color}><title>{String(x[labelKey]||"")}: {money(Number(x[valueKey]||0))}{x.orders?` · ${x.orders} orders`:""}</title></circle>)}</svg><div>{rows.map((x,i)=><span key={i}>{String(x[labelKey]||"")}</span>)}</div></div>
}
function TenderValue({row,tender}:{row:any;tender:"card"|"cash"|"talabat"|"deliveroo"|"keeta"}) {
  const value=Number(row[tender]||0);
  if(!value)return <span className="zeroDash">—</span>;
  const aggregator=["Talabat","Deliveroo","Keeta"].includes(row.orderType);
  const invalid=aggregator?tender.toLowerCase()!==row.orderType.toLowerCase():["talabat","deliveroo","keeta"].includes(tender);
  return <span className={invalid?"invalidTender":""}>{money(value)}{invalid&&<small>Unexpected channel</small>}</span>
}
type GridCol={key:string;label:string;numeric?:boolean;defaultVisible?:boolean;channel?:boolean;value?:(r:any)=>string|number;render?:(r:any)=>React.ReactNode};
type MixMetric="orders"|"inclVat"|"exVat"|"vat";
function MetricSelector({value,onChange}:{value:MixMetric;onChange:(v:MixMetric)=>void}) {
  const choices:Array<[MixMetric,string]>=[["inclVat","Revenue incl. VAT"],["exVat","Revenue excl. VAT"],["vat","VAT"],["orders","Orders"]];
  return <div className="metricSelector" aria-label="Table measure">{choices.map(([id,label])=><button type="button" key={id} className={value===id?"active":""} onClick={()=>onChange(id)}>{label}</button>)}</div>
}
function DataGrid({id,rows,columns,totals,onRowClick,selectedKey,renderExpanded,title,description,toolbarExtra,rowClassName}:{id:string;rows:any[];columns:GridCol[];totals?:Record<string,React.ReactNode>;onRowClick?:(r:any)=>void;selectedKey?:string;renderExpanded?:(r:any)=>React.ReactNode;title?:string;description?:string;toolbarExtra?:React.ReactNode;rowClassName?:(r:any)=>string}) {
  const [order,setOrder]=useState(columns.map(c=>c.key));
  const [visible,setVisible]=useState<Record<string,boolean>>(()=>Object.fromEntries(columns.map(c=>[c.key,c.defaultVisible!==false])));
  const [sort,setSort]=useState<{key:string;dir:1|-1}|null>(null);
  const [expanded,setExpanded]=useState<string|null>(null);
  const [chooser,setChooser]=useState(false);
  const colMap=useMemo(()=>new Map(columns.map(c=>[c.key,c])),[columns]);
  const active=order.filter(k=>visible[k]).map(k=>colMap.get(k)!).filter(Boolean);
  const sorted=useMemo(()=>sort?[...rows].sort((a,b)=>{const c=colMap.get(sort.key);const av=c?.value?c.value(a):a[sort.key];const bv=c?.value?c.value(b):b[sort.key];return (typeof av==="number"&&typeof bv==="number"?av-bv:String(av??"").localeCompare(String(bv??"")))*sort.dir}):rows,[rows,sort,colMap]);
  const cycleSort=(key:string)=>setSort(s=>s?.key!==key?{key,dir:1}:s.dir===1?{key,dir:-1}:null);
  const move=(from:string,to:string)=>setOrder(prev=>{const n=prev.filter(x=>x!==from);n.splice(n.indexOf(to),0,from);return n});
  const exportExcel=()=>{const header=active.map(c=>c.label);const body=sorted.map(r=>active.map(c=>{const v=c.value?c.value(r):r[c.key];if(isZeroDisplay(v))return "";return typeof v==="number"||typeof v==="string"?v:String(v??"")}));if(totals)body.push(active.map(c=>{const v=totals[c.key];if(isZeroDisplay(v))return "";return typeof v==="number"||typeof v==="string"?v:""}));const ws=XLSX.utils.aoa_to_sheet([header,...body]);ws["!cols"]=header.map((h,i)=>({wch:Math.min(40,Math.max(h.length+2,...body.slice(0,200).map(r=>String(r[i]??"").length+2)))}));const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Report");XLSX.writeFile(wb,`${id.replace(/[^a-z0-9]+/gi,"-")}.xlsx`)};
  const tableTitle=title||id.split("-").map(x=>x[0]?.toUpperCase()+x.slice(1)).join(" ");
  return <div className={`gridShell dataGrid-${id}`}><div className="gridTools"><div className="gridIdentity"><strong>{tableTitle}</strong>{description&&<small>{description}</small>}</div>{toolbarExtra}<span>{num(rows.length)} rows</span><button type="button" onClick={exportExcel}>Export Excel ↓</button><button type="button" onClick={e=>{e.stopPropagation();setChooser(v=>!v)}}>Columns ▾</button>{chooser&&<div className="columnChooser" role="group" aria-label="Visible table columns" onClick={e=>e.stopPropagation()}><strong>Show or hide columns</strong>{columns.map(c=><label key={c.key} onClick={e=>e.stopPropagation()}><input type="checkbox" checked={visible[c.key]!==false} onClick={e=>e.stopPropagation()} onChange={e=>setVisible(v=>({...v,[c.key]:e.target.checked}))}/><span>{c.label}</span></label>)}</div>}</div>
    <div className="tablePanel smartGrid"><table><thead><tr>{active.map(c=><th key={c.key} draggable onDragStart={e=>e.dataTransfer.setData("text/plain",c.key)} onDragOver={e=>e.preventDefault()} onDrop={e=>move(e.dataTransfer.getData("text/plain"),c.key)} onClick={()=>cycleSort(c.key)} className={`${c.numeric?"numeric":""} ${isChannelValue(c.label)?`channelColumnHead channel-${channelKey(c.label)}`:""}`}>{c.label}<i>{sort?.key===c.key?(sort.dir===1?" ↑":" ↓"):" ↔"}</i></th>)}</tr></thead>
      <tbody>{sorted.map((r,i)=>{const rowKey=String(r.id||r.rowId||r.name||r.date||r.week||i);const isExpanded=expanded===rowKey;return [
        <tr key={rowKey} onClick={()=>{if(renderExpanded)setExpanded(isExpanded?null:rowKey);onRowClick?.(r)}} className={`${(selectedKey&&(r.billNo===selectedKey||r.id===selectedKey))?"selected":""} ${renderExpanded?"expandableRow":""} ${(r.day==="Saturday"||r.day==="Sunday")?"weekendRow":""} ${rowClassName?.(r)||""}`}>{active.map((c,j)=>{const raw=c.value?c.value(r):r[c.key];return <td key={c.key} className={c.numeric?"numeric":""}>{j===0&&renderExpanded&&<span className="expandGlyph">{isExpanded?"−":"+"}</span>}{isZeroDisplay(raw)?<span className="zeroValue">—</span>:c.render?c.render(r):(isChannelColumn(c)&&isChannelValue(raw)?<ChannelBadge name={String(raw)}/>:String(raw??"—"))}</td>})}</tr>,
        isExpanded&&renderExpanded?<tr key={`${rowKey}-detail`} className="expandedDetailRow"><td colSpan={active.length}>{renderExpanded(r)}</td></tr>:null
      ]})}</tbody>
      {totals&&<tfoot><tr>{active.map(c=>{const value=totals[c.key];return <td key={c.key} className={c.numeric?"numeric":""}>{isZeroDisplay(value)?"—":value??""}</td>})}</tr></tfoot>}</table></div></div>
}

function MixDrill({invoices,paymentExceptions=[],metric="inclVat"}:{invoices:Invoice[];paymentExceptions?:any[];metric?:MixMetric}) {
  const exceptionMap=new Map(paymentExceptions.map(x=>[x.billNo,x]));
  const tenders=["Card","Cash","TALABAT","Deliveroo","Keeta"];
  const matrix:any={};
  invoices.forEach(x=>{
    const split=exceptionMap.get(x.billNo);
    const allocations=split?["card","cash","talabat","deliveroo","keeta"].map(k=>({name:k[0].toUpperCase()+k.slice(1),value:Number(split[k]||0)})).filter(y=>y.value>0):[{name:x.paymentMode||"Unclassified",value:x.total}];
    const row=matrix[x.orderType]??={name:x.orderType,orders:0,inclVat:0,vat:0,charges:0,cells:{}};
    row.orders++;row.inclVat+=x.total;row.vat+=x.vat;row.charges+=x.charges;
    allocations.forEach(y=>{const name=y.name.toUpperCase()==="TALABAT"?"TALABAT":y.name;const share=x.total?y.value/x.total:0;row.cells[name]??={orders:0,inclVat:0,vat:0};row.cells[name].orders++;row.cells[name].inclVat+=y.value;row.cells[name].vat+=x.vat*share});
  });
  const value=(x:any)=>metric==="orders"?Number(x.orders||0):metric==="vat"?Number(x.vat||0):metric==="exVat"?Number(x.inclVat||0)-Number(x.vat||0):Number(x.inclVat||0);
  const format=(v:number)=>metric==="orders"?num(v):money(v);
  const rows=Object.values(matrix).sort((a:any,b:any)=>value(b)-value(a)) as any[];
  const totals=(tender?:string)=>rows.reduce((a,r)=>a+value(tender?(r.cells[tender]||{}):r),0);
  const invalid=(orderType:string,tender:string)=>{const aggregators=["Talabat","Deliveroo","Keeta","Noon Food","Careem","Smiles"];return aggregators.includes(orderType)?tender.toLowerCase()!==orderType.toLowerCase():aggregators.some(x=>x.toLowerCase()===tender.toLowerCase())};
  return <div className="matrixDrill"><div className="matrixDrillHead"><div><h4>Order type × payment type</h4><p>Rows are supply channels; columns are tender allocations. Charges are shown against their originating order type. Red cells require mapping review.</p></div><Badge tone="good">{metric==="orders"?"ORDERS":metric==="vat"?"VAT":metric==="exVat"?"EXCL. VAT":"INCL. VAT"}</Badge></div><div className="matrixDrillScroll"><table><thead><tr><th>Order type</th>{tenders.map(t=><th className={`channelHead channel-${t.toLowerCase().replace(/\s+/g,"-")}`} key={t}>{t}</th>)}<th>Charges</th><th>Total</th><th>Dominant tender</th></tr></thead><tbody>{rows.map(r=>{const dominant=[...tenders].sort((a,b)=>value(r.cells[b]||{})-value(r.cells[a]||{}))[0];return <tr key={r.name}><td><ChannelBadge name={r.name}/><small>{num(r.orders)} invoices</small></td>{tenders.map(t=>{const amount=value(r.cells[t]||{}),bad=amount>0&&invalid(r.name,t);return <td className={bad?"matrixMismatch":""} key={t}>{amount?format(amount):"—"}{bad&&<small>Unexpected mapping</small>}</td>})}<td>{r.charges?money(r.charges):"—"}</td><td><strong>{format(value(r))}</strong></td><td><ChannelBadge name={dominant}/></td></tr>})}</tbody><tfoot><tr><td>TOTAL</td>{tenders.map(t=><td key={t}>{totals(t)?format(totals(t)):"—"}</td>)}<td>{money(rows.reduce((a,r)=>a+r.charges,0))}</td><td>{format(totals())}</td><td>—</td></tr></tfoot></table></div></div>
}

export default function Home() {
  const [data,setData]=useState<AuditData|null>(null);
  const [paymentRecon,setPaymentRecon]=useState<PaymentReconciliation|null>(null);
  const [separateAudit,setSeparateAudit]=useState<SeparateAudit|null>(null);
  const [kotAudit,setKotAudit]=useState<KotAudit|null>(null);
  const [businessInsights,setBusinessInsights]=useState<BusinessInsights|null>(null);
  const [tab,setTab]=useState("overview");
  const [query,setQuery]=useState("");
  const [orderType,setOrderType]=useState("All");
  const [priceStatus,setPriceStatus]=useState("All");
  const [controlFilter,setControlFilter]=useState("All");
  const [discountFilter,setDiscountFilter]=useState("All");
  const [vatFilter,setVatFilter]=useState("All");
  const [dateFilter,setDateFilter]=useState("All");
  const [selectedBill,setSelectedBill]=useState<string>("2990");
  const [reportSearch,setReportSearch]=useState("");
  const [selectedReport,setSelectedReport]=useState("Sales / Z Summary");
  const [selectedSource,setSelectedSource]=useState(sourceAudits[0].name);
  const [separateStatus,setSeparateStatus]=useState("All");
  const [separateIssue,setSeparateIssue]=useState("All");
  const [uploadError,setUploadError]=useState("");
  const [selectedSeparateBill,setSelectedSeparateBill]=useState("");
  const [costs,setCosts]=useState<Record<string,number>>({});
  const [targetFoodCost,setTargetFoodCost]=useState(30);
  const [kotQuery,setKotQuery]=useState("");
  const deferredKotQuery=useDeferredValue(kotQuery);
  const [kotStatus,setKotStatus]=useState("All");
  const [kotUser,setKotUser]=useState("All");
  const [kotDate,setKotDate]=useState("All");
  const [kotPage,setKotPage]=useState(1);
  const [selectedKotBill,setSelectedKotBill]=useState("2990");
  const [discountQuery,setDiscountQuery]=useState("");
  const [discountLevel,setDiscountLevel]=useState("All");
  const [discountPage,setDiscountPage]=useState(1);
  const [mixMetric,setMixMetric]=useState<MixMetric>("inclVat");
  const [itemRankQuery,setItemRankQuery]=useState("");
  const [sidebarCollapsed,setSidebarCollapsed]=useState(false);
  const [clock,setClock]=useState(new Date());
  useEffect(()=>{fetch("/data/audit-data.json").then(r=>r.json()).then(setData)},[]);
  useEffect(()=>{fetch("/data/payment-reconciliation.json").then(r=>r.json()).then(setPaymentRecon)},[]);
  useEffect(()=>{fetch("/data/separate-sales-audit.json").then(r=>r.json()).then(setSeparateAudit)},[]);
  useEffect(()=>{fetch("/data/item-kot-audit.json").then(r=>r.json()).then(setKotAudit)},[]);
  useEffect(()=>{fetch("/data/business-insights.json").then(r=>r.json()).then(setBusinessInsights)},[]);
  useEffect(()=>{const timer=setInterval(()=>setClock(new Date()),30000);return()=>clearInterval(timer)},[]);
  const handleSalesUpload=async(file?:File)=>{
    if(!file)return;setUploadError("");
    try{const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellDates:false});const rows=XLSX.utils.sheet_to_json<any[]>(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:null});const audited=auditUploadedSalesRows(rows,file.name);setSeparateAudit(audited);setSelectedSeparateBill("");setSeparateStatus("All");setSeparateIssue("All");setQuery("");}
    catch(e:any){setUploadError(e?.message||"Could not audit this workbook.")}
  };
  const downloadSeparateAudit=()=>{
    if(!separateAudit)return;
    const wb=XLSX.utils.book_new();
    const summaryRows=[["Audit Any Sales Report",separateAudit.summary.file],["Period",`${separateAudit.summary.periodFrom} – ${separateAudit.summary.periodTo}`],[],["Metric","Value"],["Fulfilled bills",separateAudit.summary.fulfilled],["Subtotal",separateAudit.summary.subtotal],["Discount",separateAudit.summary.discount],["VAT",separateAudit.summary.vat],["VAT control",separateAudit.summary.vatControl],["Charges",separateAudit.summary.charges],["Tax on charges",separateAudit.summary.taxOnCharges],["Revenue",separateAudit.summary.total],["Review bills",separateAudit.summary.reviewBills]];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(summaryRows),"Audit Summary");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(separateAudit.invoices),"Invoice Audit");
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(separateAudit.orderPayment),"Order Payment VAT");
    XLSX.writeFile(wb,`TMBill_Sales_Report_Audit_${new Date().toISOString().slice(0,10)}.xlsx`);
  };
  useEffect(()=>{try{setCosts(JSON.parse(localStorage.getItem("tmbill-food-costs")||"{}"))}catch{}},[]);
  useEffect(()=>{if(Object.keys(costs).length)localStorage.setItem("tmbill-food-costs",JSON.stringify(costs))},[costs]);
  const invoiceMap=useMemo(()=>new Map(data?.invoices.map(x=>[x.id,x])||[]),[data]);
  const filteredInvoices=useMemo(()=>{
    if(!data)return[];
    const q=query.toLowerCase().trim();
    return data.invoices.filter(x=>(dateFilter==="All"||x.date===dateFilter)&&(orderType==="All"||x.orderType===orderType)&&(controlFilter==="All"||x.controlStatus===controlFilter)&&(discountFilter==="All"||x.discountAssessment===discountFilter)&&(vatFilter==="All"||x.vatAssessment===vatFilter)&&(!q||x.billNo.includes(q)||x.id.toLowerCase().includes(q)||x.user.toLowerCase().includes(q)||x.paymentMode.toLowerCase().includes(q)));
  },[data,query,orderType,controlFilter,discountFilter,vatFilter,dateFilter]);
  const filteredItems=useMemo(()=>{
    if(!data)return[];
    const q=query.toLowerCase().trim();
    return data.items.filter(x=>(orderType==="All"||x.orderType===orderType)&&(priceStatus==="All"||x.priceStatus===priceStatus)&&(!q||x.billNo.includes(q)||x.name.toLowerCase().includes(q)||x.itemId.toLowerCase().includes(q)));
  },[data,query,orderType,priceStatus]);
  const selected=useMemo(()=>data?.invoices.find(x=>x.billNo===selectedBill)||null,[data,selectedBill]);
  const selectedLines=useMemo(()=>selected?data?.items.filter(x=>x.invoiceId===selected.id)||[]:[],[data,selected]);
  const selectedTurnaround=selected?businessInsights?.invoiceTurnaround[selected.id]||null:null;
  const separateRows=useMemo(()=>separateAudit?.invoices.filter(x=>(separateStatus==="All"||x.assessment===separateStatus)&&(separateIssue==="All"||x.issues.includes(separateIssue))&&(!query||x.billNo.includes(query)||x.id.toLowerCase().includes(query.toLowerCase())||x.orderType.toLowerCase().includes(query.toLowerCase())))||[],[separateAudit,separateStatus,separateIssue,query]);
  const separateSelected=useMemo(()=>separateAudit?.invoices.find(x=>x.billNo===selectedSeparateBill)||null,[separateAudit,selectedSeparateBill]);
  const separateMatrix=useMemo(()=>{
    if(!separateAudit)return[];
    const m=new Map<string,any>();
    for(const x of separateAudit.invoices.filter(x=>x.status.toLowerCase()==="fulfilled")){
      const z=m.get(x.orderType)||{orderType:x.orderType,orders:0,card:0,cash:0,talabat:0,deliveroo:0,keeta:0,correctSales:0,invoiceVat:0,taxReportCaptured:0};
      z.orders++;z.correctSales+=x.total;z.invoiceVat+=x.vat;
      const key=String(x.paymentMode||"").toLowerCase();
      if(["card","cash","talabat","deliveroo","keeta"].includes(key)){z[key]+=x.total;z.taxReportCaptured+=x.total}
      m.set(x.orderType,z);
    }
    return [...m.values()].map(x=>Object.fromEntries(Object.entries(x).map(([k,v])=>[k,typeof v==="number"?+v.toFixed(2):v])));
  },[separateAudit]);
  const filteredKotEvents=useMemo(()=>{
    if(!kotAudit)return[];
    const q=deferredKotQuery.trim().toLowerCase();
    return kotAudit.events.filter(x=>
      (kotStatus==="All"||x.status===kotStatus)&&
      (kotUser==="All"||x.user===kotUser)&&
      (kotDate==="All"||x.date===kotDate)&&
      (!q||[x.billNo,x.orderId,x.kotId,x.kotNo,x.item,x.user,x.reason,x.table].some(v=>String(v||"").toLowerCase().includes(q)))
    );
  },[kotAudit,deferredKotQuery,kotStatus,kotUser,kotDate]);
  const kotPageSize=75;
  const kotPageCount=Math.max(1,Math.ceil(filteredKotEvents.length/kotPageSize));
  const pagedKotEvents=useMemo(()=>filteredKotEvents.slice((kotPage-1)*kotPageSize,kotPage*kotPageSize),[filteredKotEvents,kotPage]);
  const selectedKotEvents=useMemo(()=>kotAudit?.events.filter(x=>x.billNo===selectedKotBill).sort((a,b)=>String(a.punchTime).localeCompare(String(b.punchTime)))||[],[kotAudit,selectedKotBill]);
  const selectedKotSummary=useMemo(()=>kotAudit?.bills.find(x=>x.billNo===selectedKotBill)||null,[kotAudit,selectedKotBill]);
  useEffect(()=>setKotPage(1),[deferredKotQuery,kotStatus,kotUser,kotDate]);
  useEffect(()=>{if(kotPage>kotPageCount)setKotPage(kotPageCount)},[kotPage,kotPageCount]);
  const filteredDiscountDetails=useMemo(()=>{
    if(!businessInsights)return[];
    const q=discountQuery.trim().toLowerCase();
    return businessInsights.discountDetails.filter(x=>(discountLevel==="All"||x.level===discountLevel)&&(!q||[x.billNo,x.orderId,x.title,x.reason,x.user,x.orderType].some(v=>String(v||"").toLowerCase().includes(q))));
  },[businessInsights,discountQuery,discountLevel]);
  const discountPageSize=100;
  const discountPageCount=Math.max(1,Math.ceil(filteredDiscountDetails.length/discountPageSize));
  const pagedDiscountDetails=useMemo(()=>filteredDiscountDetails.slice((discountPage-1)*discountPageSize,discountPage*discountPageSize),[filteredDiscountDetails,discountPage]);
  useEffect(()=>setDiscountPage(1),[discountQuery,discountLevel]);
  const menuRows=useMemo(()=>{
    if(!data)return[];
    const avgQty=data.topItems.reduce((a,x)=>a+x.qty,0)/Math.max(1,data.topItems.length);
    const enriched=data.topItems.map(x=>{const price=x.qty?x.net/x.qty:0,cost=costs[x.name]||0,margin=price-cost;return {...x,price,cost,margin,totalMargin:x.net-cost*x.qty}});
    const avgMargin=enriched.filter(x=>x.cost>0).reduce((a,x)=>a+x.margin,0)/Math.max(1,enriched.filter(x=>x.cost>0).length);
    return enriched.map(x=>({...x,classification:x.qty>=avgQty?(x.margin>=avgMargin?"Star":"Plowhorse"):(x.margin>=avgMargin?"Puzzle":"Dog"),idealPrice:x.cost/(targetFoodCost/100||.3)}));
  },[data,costs,targetFoodCost]);
  if(!data) return <main className="loading"><div className="loader"/><h1>Preparing the TMBill audit workspace</h1><p>Linking invoices, items, prices and controls…</p></main>;
  const s=data.summary as Record<string,number|string>;
  const orderTotals=[
    ["Dine In",35247.7,684],["Pickup",20262.85,488],["Talabat",42549.7,577],["Delivery",8657.5,52],["Keeta",1455,15]
  ] as const;
  const paymentTotals=[
    ["Card",53321.15,936],["Cash",10729.9,301],["TALABAT",42601.1,578],["Keeta",1520.6,16]
  ] as const;
  const orderTypeNames=["Dine In","Pickup","Delivery","Talabat","Keeta"];
  const weekdayNames=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const invoiceDay=(x:Invoice)=>{const [d,m,y]=x.date.split("-").map(Number);return new Date(y,m-1,d)};
  const weekdayOrderMix=weekdayNames.map(day=>{
    const invoices=data.invoices.filter(x=>x.status==="Fulfilled"&&weekdayNames[invoiceDay(x).getDay()]===day);
    const row:any={day,orders:invoices.length,totalInclVat:invoices.reduce((a,x)=>a+x.total,0),totalVat:invoices.reduce((a,x)=>a+x.vat,0)};
    orderTypeNames.forEach(type=>{const lines=invoices.filter(x=>x.orderType===type);row[`${type}InclVat`]=lines.reduce((a,x)=>a+x.total,0);row[`${type}Vat`]=lines.reduce((a,x)=>a+x.vat,0);row[`${type}Orders`]=lines.length});
    const metric=(type:string)=>mixMetric==="orders"?row[`${type}Orders`]:mixMetric==="vat"?row[`${type}Vat`]:mixMetric==="exVat"?row[`${type}InclVat`]-row[`${type}Vat`]:row[`${type}InclVat`];
    orderTypeNames.forEach(type=>row[type]=metric(type));
    row.total=mixMetric==="orders"?row.orders:mixMetric==="vat"?row.totalVat:mixMetric==="exVat"?row.totalInclVat-row.totalVat:row.totalInclVat;
    const ranked=orderTypeNames.map(type=>({type,value:metric(type)})).sort((a,b)=>b.value-a.value);
    row.dominant=ranked[0]?.type||"—";row.weakest=[...ranked].reverse().find(x=>x.value>0)?.type||"—";
    return row;
  });
  const mixValue=(value:number)=>mixMetric==="orders"?num(value):money(value);
  const mixLabel=mixMetric==="orders"?"Order count":mixMetric==="vat"?"VAT":mixMetric==="exVat"?"Revenue excl. VAT":"Revenue incl. VAT";
  const biMetric=(x:any,average=false)=>{const divisor=average?Math.max(1,Number(x.days||1)):1;return (mixMetric==="orders"?Number(x.orders||0):mixMetric==="vat"?Number(x.vat||0):mixMetric==="exVat"?Number(x.revenue||0)-Number(x.vat||0):Number(x.revenue||0))/divisor};
  const weekdayPerformanceRows=weekdayNames.map(day=>businessInsights?.weekday.find(x=>x.day===day)).filter(Boolean).map((x:any)=>({...x,displayValue:biMetric(x,true)}));
  const metricFrom=(orders:number,inclVat:number,vat:number)=>mixMetric==="orders"?orders:mixMetric==="vat"?vat:mixMetric==="exVat"?inclVat-vat:inclVat;
  const orderMixRows=orderTypeNames.map(name=>{const lines=data.invoices.filter(x=>x.status==="Fulfilled"&&x.orderType===name);const inclVat=lines.reduce((a,x)=>a+x.total,0),vat=lines.reduce((a,x)=>a+x.vat,0),charges=lines.reduce((a,x)=>a+x.charges,0);return {name,orders:lines.length,inclVat,vat,charges,netItemsInclVat:inclVat-charges,value:metricFrom(lines.length,inclVat,vat)}}).filter(x=>x.orders);
  const paymentExceptionMap=new Map((paymentRecon?.exceptions||[]).map((x:any)=>[x.billNo,x]));
  const channelMismatchBills=new Set((paymentRecon?.exceptions||[]).filter((x:any)=>x.channelMismatch).map((x:any)=>x.billNo));
  const hasChannelMismatch=(invoices:Invoice[])=>invoices.some(x=>channelMismatchBills.has(x.billNo));
  const invoicesForTender=(name:string)=>data.invoices.filter(x=>{if(x.status!=="Fulfilled")return false;const split:any=paymentExceptionMap.get(x.billNo);if(split)return Number(split[name.toLowerCase().replace(" ","")]||0)>0;return x.paymentMode.toLowerCase().split(",").map(v=>v.trim()).includes(name.toLowerCase())});
  const paymentMixRows=Object.values(data.invoices.filter(x=>x.status==="Fulfilled").reduce((a:any,x)=>{
    const split:any=paymentExceptionMap.get(x.billNo);
    const allocations=split?["card","cash","talabat","deliveroo","keeta"].map(k=>({name:k[0].toUpperCase()+k.slice(1),value:Number(split[k]||0)})).filter(y=>y.value>0):[{name:x.paymentMode||"Unclassified",value:x.total}];
    allocations.forEach(y=>{const share=x.total?y.value/x.total:0;a[y.name]??={name:y.name,orders:0,inclVat:0,vat:0};a[y.name].orders++;a[y.name].inclVat+=y.value;a[y.name].vat+=x.vat*share});
    return a;
  },{})).map((x:any)=>({...x,value:metricFrom(x.orders,x.inclVat,x.vat)})).sort((a:any,b:any)=>b.value-a.value) as any[];
  const orderMetricTotal=orderMixRows.reduce((a,x)=>a+x.value,0),paymentMetricTotal=paymentMixRows.reduce((a,x)=>a+x.value,0);
  const weekPartRows=[{name:"Weekdays",days:"Monday–Friday",dayCount:22,test:(x:Invoice)=>![0,6].includes(invoiceDay(x).getDay())},{name:"Weekend",days:"Saturday–Sunday",dayCount:8,test:(x:Invoice)=>[0,6].includes(invoiceDay(x).getDay())}].map(group=>{const lines=data.invoices.filter(x=>x.status==="Fulfilled"&&group.test(x));const revenue=lines.reduce((a,x)=>a+x.total,0),vat=lines.reduce((a,x)=>a+x.vat,0);return {name:group.name,days:group.days,dayCount:group.dayCount,orders:lines.length,revenue,vat,revenueExVat:revenue-vat,averageDaily:revenue/group.dayCount,averageCheck:revenue/Math.max(1,lines.length),displayValue:metricFrom(lines.length,revenue,vat)}});
  const channelGroupRows=[{name:"Own Orders",types:["Dine In","Pickup","Delivery"],description:"Dine In · Pickup · Delivery"},{name:"Aggregators",types:["Talabat","Noonfood","Deliveroo","Careem","Smiles","Keeta"],description:"Online food aggregators"}].map(group=>{const lines=data.invoices.filter(x=>x.status==="Fulfilled"&&group.types.includes(x.orderType));const revenue=lines.reduce((a,x)=>a+x.total,0),vat=lines.reduce((a,x)=>a+x.vat,0),discount=lines.reduce((a,x)=>a+x.discount,0),charges=lines.reduce((a,x)=>a+x.charges,0),guests=lines.reduce((a,x)=>a+Number(x.guests||0),0);return {...group,orders:lines.length,guests,revenue,vat,discount,charges,revenueExVat:revenue-vat,averageCheck:revenue/Math.max(1,lines.length),mix:revenue/+s.grossSales*100,displayValue:metricFrom(lines.length,revenue,vat)}});
  const zOrderRows=orderTypeNames.map(name=>{const fulfilled=data.invoices.filter(x=>x.status==="Fulfilled"&&x.orderType===name),cancelled=data.invoices.filter(x=>x.status!=="Fulfilled"&&x.orderType===name);const revenue=fulfilled.reduce((a,x)=>a+x.total,0),vat=fulfilled.reduce((a,x)=>a+x.vat,0),discount=fulfilled.reduce((a,x)=>a+x.discount,0),charges=fulfilled.reduce((a,x)=>a+x.charges,0),gross=fulfilled.reduce((a,x)=>a+x.canonicalGrossBeforeDiscount,0),complimentary=fulfilled.reduce((a,x)=>a+x.canonicalItemDiscount,0),guests=fulfilled.reduce((a,x)=>a+Number(x.guests||0),0);const netItemsInclVat=revenue-charges;return {name,orders:fulfilled.length,guests,gross,revenue,discount,complimentary,charges,netItemsInclVat,vat,net:netItemsInclVat-vat,cancelled:cancelled.length,cancelledValue:cancelled.reduce((a,x)=>a+x.total,0),flag:hasChannelMismatch(fulfilled)}});
  const invoiceHour=(x:Invoice)=>{const match=String(x.time||"").match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);if(!match)return 0;let hour=Number(match[1])%12;if(match[3].toUpperCase()==="PM")hour+=12;return hour+Number(match[2])/60};
  const mealDefinitions=[["Midnight Sales","12:00 AM – 06:00 AM",0,6],["Breakfast","06:00 AM – 11:00 AM",6,11],["Lunch","11:00 AM – 04:00 PM",11,16],["Snacks","04:00 PM – 07:00 PM",16,19],["Dinner","07:00 PM – 12:00 AM",19,24]] as const;
  const mealRows=mealDefinitions.map(([name,timeSlot,from,to])=>{const lines=data.invoices.filter(x=>x.status==="Fulfilled"&&invoiceHour(x)>=from&&invoiceHour(x)<to);const revenue=lines.reduce((a,x)=>a+x.total,0),vat=lines.reduce((a,x)=>a+x.vat,0),charges=lines.reduce((a,x)=>a+x.charges,0),discount=lines.reduce((a,x)=>a+x.discount,0);return {name,timeSlot,from,to,orders:lines.length,revenue,vat,charges,discount,averageCheck:revenue/Math.max(1,lines.length),mix:revenue/+s.grossSales*100,displayValue:metricFrom(lines.length,revenue,vat),invoices:lines}});
  const searchedRankItems=data.topItems.filter(x=>x.name.toLowerCase().includes(itemRankQuery.trim().toLowerCase()));
  const averageCards=[
    {label:"Average per calendar day",value:money(+s.grossSales/30),sub:"30 June business dates",tab:"overview"},
    {label:"Average weekday day",value:money(weekPartRows[0].averageDaily),sub:"Monday–Friday · 22 days",tab:"overview"},
    {label:"Average weekend day",value:money(weekPartRows[1].averageDaily),sub:"Saturday–Sunday · 8 days",tab:"overview"},
    {label:"Average per 7-day week",value:money(+s.grossSales/(30/7)),sub:"Normalized weekly run rate",tab:"overview"},
    {label:"Average cheque",value:money(+s.averageCheck),sub:"Revenue per fulfilled bill",tab:"bills"},
    {label:"Average per person",value:money(+s.averagePerPerson),sub:`${num(+s.totalGuests)} recorded guests`,tab:"bills"},
    {label:"Average items per bill",value:(+s.totalItemQuantity/+s.fulfilledInvoices).toFixed(2),sub:`${num(+s.totalItemQuantity)} items / ${num(+s.fulfilledInvoices)} bills`,tab:"items"},
    {label:"Average turnaround",value:minutesLabel(businessInsights?.summary.overallAverageMinutes),sub:`Median ${minutesLabel(businessInsights?.summary.overallMedianMinutes)}`,tab:"operations"},
  ];
  const invoicesForDay=(date:string)=>data.invoices.filter(x=>x.status==="Fulfilled"&&`${String(invoiceDay(x).getFullYear())}-${String(invoiceDay(x).getMonth()+1).padStart(2,"0")}-${String(invoiceDay(x).getDate()).padStart(2,"0")}`===date);
  const invoicesForWeek=(week:string)=>{const index=businessInsights?.weekly.findIndex(x=>x.week===week)??-1;return index<0?[]:data.invoices.filter(x=>{const d=invoiceDay(x).getDate();return x.status==="Fulfilled"&&d>=index*7+1&&d<=Math.min(30,index*7+7)})};
  return <div className={`app ${sidebarCollapsed?"navCollapsed":""}`}>
    <aside className="sidebar">
      <div className="brand"><div className="brandLogoCrop"><img src="/tmbill-logo-source.png" alt="TMBill Technology LLC"/></div><div><b>TMBill Audit</b><small>Revenue intelligence</small></div><button className="navToggle" onClick={()=>setSidebarCollapsed(v=>!v)} title={sidebarCollapsed?"Expand navigation":"Collapse navigation"}>{sidebarCollapsed?"›":"‹"}</button></div>
      <div className="period"><span>Review period</span><b>June 2026</b><small>The SS · Morbido Express</small></div>
      <nav>{tabs.map(([id,label])=><button key={id} title={label} data-label={label} className={tab===id?"active":""} onClick={()=>{setTab(id);if(sidebarCollapsed){setSidebarCollapsed(false);window.setTimeout(()=>setSidebarCollapsed(true),1200)}}}><i>{tabIcons[id]}</i><span>{label}</span>{id==="bills"&&<em>1</em>}{id==="items"&&<em>915</em>}</button>)}</nav>
      <div className="sideControl"><span>Overall reconciliation</span><strong>3 confirmed defects</strong><small>Core invoice ledger passes</small></div>
    </aside>
    <main className="main">
      <header className="topbar">
        <div><div className="eyebrow">THE SS · MORBIDO EXPRESS RESTAURANT</div><h1>{tabs.find(x=>x[0]===tab)?.[1]}</h1></div>
        <div className="topActions"><span className="liveClock"><b>{clock.toLocaleTimeString("en-AE",{hour:"2-digit",minute:"2-digit"})}</b><small>Dubai time</small></span><button className="iconBtn" title="Notifications">●<em>3</em></button><span className="sync"><i/> Data frozen at 30 Jun 2026, 11:58 PM</span><button className="helpBtn" onClick={()=>setTab("guide")}>Metric guide</button></div>
      </header>

      {tab==="overview"&&<>
        <section className="welcomeBanner"><div className="welcomeLogoCrop"><img src="/tmbill-logo-source.png" alt="TMBill"/></div><div><span>HELLO TMBILL TEAM</span><h2>Welcome to the TMBill Audit Dashboard</h2><p>Every reconciled day makes reporting clearer, operations stronger, and decisions safer. We are getting better—one controlled bill at a time.</p></div><button onClick={()=>setTab("guide")}>Start guided review →</button></section>
        <section className="hero">
          <div><Badge tone="good">Invoice ledger reconciled</Badge><h2>Sales are right. Some dimensions are not.</h2><p>The fulfilled invoice ledger ties across reports. The red controls below identify where item, payment and category calculations break after invoice settlement.</p></div>
          <div className="heroScore"><b>97%</b><span>financial control confidence</span><small>3 exceptions require correction</small></div>
        </section>
        <section className="averageSequence"><div className="sequenceHead"><div><span>MANAGEMENT AVERAGES</span><h3>June operating rhythm at a glance</h3></div><small>Time period → transaction → customer experience</small></div><div>{averageCards.map(x=><button key={x.label} onClick={()=>setTab(x.tab)}><span>{x.label}</span><b>{x.value}</b><small>{x.sub}</small></button>)}</div></section>
        <section className="kpiGrid">
          <Kpi onClick={()=>setTab("items")} label="Gross Item Sales" value={money(+s.correctedGrossBeforeDiscount)} sub={`${num(+s.totalItemQuantity)} items before any reductions`}/>
          <Kpi onClick={()=>setTab("discounts")} label="Discounts" value={money(+s.canonicalOrderDiscount)} sub="Order and promotional reductions"/>
          <Kpi onClick={()=>{setSelectedBill("2990");setTab("bills")}} label="Complimentary" value={money(+s.canonicalItemDiscount)} sub="Fattoush #2990 · reason, actor and time retained"/>
          <Kpi onClick={()=>setTab("discounts")} label="Total Reductions" value={money(+s.canonicalTotalDiscount)} sub="Discounts + Complimentary"/>
          <Kpi onClick={()=>setTab("crosscheck")} label="Net Item Sales incl. VAT" value={money(+s.taxableGrossInclVat)} sub="Gross Item Sales − Total Reductions"/>
          <Kpi onClick={()=>setTab("vat")} label="VAT" value={money(+s.vat)} sub={`Included tax disclosure · control ${money(5124.93)}`}/>
          <Kpi onClick={()=>setTab("vat")} label="Net Item Sales excl. VAT" value={money(+s.taxableGrossInclVat-+s.vat)} sub="Net Item Sales incl. VAT − VAT"/>
          <Kpi onClick={()=>setTab("crosscheck")} label="Charges incl. VAT" value={money(+s.charges)} sub="Must be classified by charge type and tax treatment"/>
          <Kpi onClick={()=>setTab("zreport")} label="Actual Sales incl. VAT" value={money(+s.grossSales)} sub="Net Item Sales incl. VAT + Charges incl. VAT"/>
          <Kpi onClick={()=>setTab("bills")} label="Bill series" value={`#${s.billFrom}–#${s.billTo}`} sub={`${num(+s.fulfilledInvoices)} fulfilled bills generated`}/>
          <Kpi onClick={()=>setTab("bills")} label="Average cheque" value={money(+s.averageCheck)} sub="Revenue per fulfilled bill"/>
          <Kpi onClick={()=>setTab("bills")} label="Average per person" value={money(+s.averagePerPerson)} sub={`${num(+s.totalGuests)} recorded guests`}/>
          {businessInsights&&<Kpi onClick={()=>document.getElementById("turnaround-insights")?.scrollIntoView({behavior:"smooth"})} label="Average turnaround" value={minutesLabel(businessInsights.summary.overallAverageMinutes)} sub={`Median ${minutesLabel(businessInsights.summary.overallMedianMinutes)} · placed to settlement`}/>}
          {businessInsights&&<Kpi onClick={()=>document.getElementById("turnaround-insights")?.scrollIntoView({behavior:"smooth"})} tone="warn" label="Turnaround coverage" value={`${num(businessInsights.summary.ordersWithTurnaround)} / ${num(+s.fulfilledInvoices)}`} sub={`${num(businessInsights.summary.ordersMissingTurnaround)} missing · mainly aggregator orders`}/>}
          <Kpi onClick={()=>setTab("payments")} tone="bad" label="Payment allocation gap" value={money(+s.paymentComponentGap)} sub="Missing secondary split-tender allocations"/>
          <Kpi onClick={()=>setTab("crosscheck")} tone="warn" label="Charge classification gap" value={money(61)} sub="All charges 581 vs DSR delivery charges 520"/>
          <Kpi onClick={()=>setTab("categories")} tone="bad" label="Category overstatement" value={money(+s.categoryOverstatement)} sub="Item export net exceeds invoice total"/>
          <Kpi onClick={()=>setTab("operations")} tone="bad" label="Cancelled KOT items" value={num(kotAudit?.summary.cancelledItems||103)} sub={`${money(kotAudit?.summary.cancelledListValue||4794)} listed value · operational leakage`}/>
          <Kpi onClick={()=>setTab("zreport")} tone="warn" label="Cancelled bills" value="13" sub={`${money(1094.50)} cancelled invoice value`}/>
        </section>
        <section className="developerControlPanel">
          <div className="salesBridgeVisual"><div className="panelHead"><div><h3>June sales bridge · one calculation for every report</h3><p>Follow the arrows from transaction-time item value to the invoice/customer total.</p></div><Badge tone="good">RECONCILES</Badge></div>
            <div className="bridgeFlow">
              <div><span>Gross Item Sales</span><b>{money(118066)}</b></div><i>−</i><div><span>Discounts</span><b>{money(10461.25)}</b></div><i>−</i><div className="complimentaryStep"><span>Complimentary</span><b>{money(13)}</b><small>#2990 Fattoush</small></div><i>=</i><div><span>Net Item Sales incl. VAT</span><b>{money(107591.75)}</b></div><i>+</i><div><span>Charges incl. VAT</span><b>{money(581)}</b></div><i>=</i><div className="actualStep"><span>Actual Sales incl. VAT</span><b>{money(108172.75)}</b></div>
            </div>
            <div className="vatDisclosure"><span>VAT included</span><b>{money(5124.92)}</b><i>→</i><span>Net Item Sales excl. VAT</span><b>{money(102466.83)}</b><small>VAT is disclosed separately; it is not deducted from Actual Sales incl. VAT.</small></div>
          </div>
          <div className="issueMap"><div className="panelHead"><div><h3>Developer root-cause map</h3><p>What users see, where the data breaks, and the required control.</p></div><Badge tone="bad">3 CORE FAILURES</Badge></div>
            <button onClick={()=>setTab("items")}><em>ITEM / CATEGORY</em><b>TMBill loses {money(13)} original gross</b><span>Zero price + discount representation breaks item/category totals.</span><strong>Persist original price + Complimentary event →</strong></button>
            <button onClick={()=>setTab("payments")}><em>PAYMENT / TAX EXPORT</em><b>{money(1181.9)} secondary tenders omitted</b><span>Invoice totals pass; tender-component columns do not.</span><strong>Explode invoice × tender allocations →</strong></button>
            <button onClick={()=>setTab("crosscheck")}><em>CHARGES</em><b>{money(61)} disappears from DSR</b><span>Invoice charges {money(581)} versus delivery-only {money(520)}.</span><strong>Normalize charge type + charge VAT →</strong></button>
          </div>
        </section>
        <section className="insightShortcuts"><div><span>JUMP TO INSIGHT</span>{[["Executive","executive-highlights"],["Turnaround","turnaround-insights"],["Supply × payment","supply-collection-insights"],["Hourly","hourly-insights"],["Weekday + order type","weekday-insights"],["Daily","daily-insights"],["Weekly","weekly-insights"],["Meal time","meal-time-insights"],["Leakage & items","protection-insights"],["Recommendations","controller-recommendations"]].map(([label,id])=><button key={id} onClick={()=>document.getElementById(id)?.scrollIntoView({behavior:"smooth",block:"start"})}>{label}</button>)}</div><div className="insightDownloads"><a href="/downloads/June_Business_Insights.pdf" download>PDF ↓</a><a href="/downloads/TMBill_June_Charge_and_Meal_Control.xlsx" download>Excel ↓</a></div></section>
        <section className="twoCol" id="executive-highlights">
          <div className="panel"><div className="panelHead"><div><h3>June executive highlights</h3><p>The strongest operating signals from the reconciled invoice ledger.</p></div><Badge tone="good">CONTROLLED</Badge></div>
            {businessInsights&&<div className="highlightList">
              <div><span>Best revenue date</span><b>{new Date(`${businessInsights.operatingHighlights.bestRevenueDay.date}T12:00:00`).toLocaleDateString("en-AE",{weekday:"long",day:"2-digit",month:"short"})}</b><strong>{money(businessInsights.operatingHighlights.bestRevenueDay.revenue)}</strong></div>
              <div><span>Highest order volume</span><b>{new Date(`${businessInsights.operatingHighlights.bestOrderDay.date}T12:00:00`).toLocaleDateString("en-AE",{weekday:"long",day:"2-digit",month:"short"})}</b><strong>{num(businessInsights.operatingHighlights.bestOrderDay.orders)} orders</strong></div>
              <div><span>Peak revenue hour</span><b>{businessInsights.operatingHighlights.peakHour.label}</b><strong>{money(businessInsights.operatingHighlights.peakHour.revenue)}</strong></div>
              <div><span>Strongest full week</span><b>{[...businessInsights.weekly].sort((a,b)=>b.revenue-a.revenue)[0]?.week}</b><strong>{money([...businessInsights.weekly].sort((a,b)=>b.revenue-a.revenue)[0]?.revenue||0)}</strong></div>
            </div>}
          </div>
          <div className="panel"><div className="panelHead"><div><h3>Issue trace</h3><p>Start with the largest control failures.</p></div></div>
            <button className="issueRow medium" onClick={()=>{setSelectedBill("2990");setTab("bills")}}><span>01</span><div><b>Complimentary control resolved</b><small>Bill 2990 · {money(13)} Fattoush linked to actor, time and reason</small></div><strong>Trace →</strong></button>
            <button className="issueRow high" onClick={()=>setTab("payments")}><span>02</span><div><b>Incomplete split payments</b><small>Tax payment components short by {money(1181.9)}</small></div><strong>Trace →</strong></button>
            <button className="issueRow medium" onClick={()=>setTab("categories")}><span>03</span><div><b>Category denominator mismatch</b><small>Percentages add to 106.29%, not 100%</small></div><strong>Trace →</strong></button>
          </div>
        </section>
        {businessInsights&&<section className="twoCol insightCharts">
          <div className="panel" id="turnaround-insights"><div className="panelHead"><div><h3>Order turnaround by order type</h3><p>Placed Time → Settlement Time · fulfilled invoices with both timestamps.</p></div><Badge tone="warn">P90 {minutesLabel(businessInsights.summary.overallP90Minutes)}</Badge></div>
            <div className="turnaroundChart">{businessInsights.turnaroundByType.map(x=><button key={x.orderType} onClick={()=>{setOrderType(x.orderType);setTab("bills")}}><div><ChannelBadge name={x.orderType}/><span>{num(x.orders)} timed orders · median {minutesLabel(x.medianMinutes)}</span></div><div className="timeTrack"><i className={`channel-${channelKey(x.orderType)}`} style={{width:`${Math.min(100,x.averageMinutes/100*100)}%`}}/><em style={{left:`${Math.min(98,x.p90Minutes/180*100)}%`}} title={`P90 ${minutesLabel(x.p90Minutes)}`}/></div><strong>{minutesLabel(x.averageMinutes)}</strong></button>)}</div>
            <div className="chartLegend"><span><i className="avg"/>Average turnaround</span><span><i className="p90"/>P90 marker</span><span>{num(businessInsights.summary.over60)} orders exceeded 60 minutes</span></div>
          </div>
          <div className="panel" id="supply-collection-insights"><div className="panelHead"><div><h3>Supply mix × collection mix</h3><p>Sortable controls with fixed totals for two dimensions of {money(108172.75)}.</p></div><Badge tone="good">RECONCILED</Badge></div>
            <div className="mixTablesRow"><div><DataGrid title="Sales by order type" description="Click a row to expand its payment matrix" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="overview-order-mix" rows={orderMixRows.map(x=>({...x,mix:x.value/Math.max(1,orderMetricTotal)*100}))} rowClassName={x=>hasChannelMismatch(data.invoices.filter(i=>i.orderType===x.name))?"paymentMismatchRow":""} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={data.invoices.filter(i=>i.status==="Fulfilled"&&i.orderType===x.name)}/>} columns={[
              {key:"name",label:"Order type"},{key:"orders",label:"Orders",numeric:true},{key:"netItemsInclVat",label:"Net item sales incl. VAT",numeric:true,render:x=>money(x.netItemsInclVat)},{key:"charges",label:"Charges incl. VAT",numeric:true,render:x=>money(x.charges)},{key:"value",label:mixLabel,numeric:true,render:x=>mixValue(x.value)},{key:"mix",label:"Mix %",numeric:true,render:x=><span className="miniMix"><i style={{width:`${x.mix}%`}}/><b>{x.mix.toFixed(1)}%</b></span>}
            ]} totals={{name:"TOTAL",orders:num(+s.fulfilledInvoices),netItemsInclVat:money(+s.taxableGrossInclVat),charges:money(+s.charges),value:mixValue(orderMetricTotal),mix:"100.0%"}}/></div>
            <div><DataGrid title="Collections by payment mode" description="Click a row to expand its order-type matrix" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="overview-payment-mix" rows={paymentMixRows.map(x=>({...x,mix:x.value/Math.max(1,paymentMetricTotal)*100}))} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={invoicesForTender(x.name)}/>} columns={[
              {key:"name",label:"Payment mode"},{key:"orders",label:"Allocations",numeric:true},{key:"value",label:mixLabel,numeric:true,render:x=>mixValue(x.value)},{key:"mix",label:"Mix %",numeric:true,render:x=><span className="miniMix payment"><i style={{width:`${x.mix}%`}}/><b>{x.mix.toFixed(1)}%</b></span>}
            ]} totals={{name:"TOTAL",orders:num(paymentMixRows.reduce((a,x)=>a+x.orders,0)),value:mixValue(paymentMetricTotal),mix:"100.0%"}}/></div></div>
            <Info title="How to read this" tone="blue">Order type explains where the sale occurred; payment mode explains how it was collected. Talabat and Keeta should usually align, while Dine In, Pickup and Delivery normally split across Card and Cash. Split tenders require multiple allocation rows.</Info>
          </div>
        </section>}
        {businessInsights&&<>
          <section className="analyticsGrid">
            <div className="panel wide" id="hourly-insights"><div className="panelHead"><div><h3>Hourly sales curve</h3><p>Fulfilled revenue by invoice hour · hover each bar for revenue and order count.</p></div><Badge tone="good">Peak {businessInsights.operatingHighlights.peakHour.label}</Badge></div>
              <div className="hourBars">{businessInsights.hourly.map(x=><div key={x.hour} title={`${x.label}: ${money(x.revenue)} · ${x.orders} orders`}><span><i style={{height:`${x.revenue/Math.max(...businessInsights.hourly.map(h=>h.revenue))*100}%`}}/></span><small>{x.hour%3===0?x.label.replace(" ",""):""}</small></div>)}</div>
            </div>
            <div className="panel wide weekdayPanel" id="weekday-insights"><div className="panelHead"><div><h3>Weekday revenue performance</h3><p>Graph on the left, sortable control table on the right.</p></div></div>
              <div className="weekdaySplit"><div><MetricLine rows={weekdayPerformanceRows} valueKey="displayValue" labelKey="day" color="#d08b22"/></div><DataGrid title="Weekday control table" description="Default Sunday–Saturday; click a header for A–Z, Z–A, then restore the operational sequence" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="weekday-performance" rows={weekdayPerformanceRows} rowClassName={x=>hasChannelMismatch(data.invoices.filter(i=>weekdayNames[invoiceDay(i).getDay()]===x.day))?"paymentMismatchRow":""} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={data.invoices.filter(i=>i.status==="Fulfilled"&&weekdayNames[invoiceDay(i).getDay()]===x.day)}/>} columns={[{key:"day",label:"Day"},{key:"days",label:"Days",numeric:true,defaultVisible:false},{key:"orders",label:"Total orders",numeric:true},{key:"displayValue",label:`Average daily ${mixLabel.toLowerCase()}`,numeric:true,render:x=>mixValue(x.displayValue)},{key:"averageCheck",label:"Avg check",numeric:true,defaultVisible:false,render:x=>money(x.averageCheck)}]} totals={{day:"TOTAL / AVG",days:"30",orders:num(+s.fulfilledInvoices),displayValue:mixValue(biMetric({orders:+s.fulfilledInvoices,revenue:+s.grossSales,vat:+s.vat,days:30},true)),averageCheck:money(+s.averageCheck)}}/></div>
            </div>
            <div className="panel wide weekdayOrderPanel"><div className="panelHead"><div><h3>Order type × weekday sales mix</h3><p>Compare each supply channel by weekday. Dominant and lowest active order types are identified for faster scheduling and channel planning.</p></div><Badge tone="good">{money(108172.75)}</Badge></div>
              <DataGrid title="Order type × weekday sales mix" description="Click a weekday to expand order type × payment type" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="weekday-order-mix" rows={weekdayOrderMix} rowClassName={x=>hasChannelMismatch(data.invoices.filter(i=>weekdayNames[invoiceDay(i).getDay()]===x.day))?"paymentMismatchRow":""} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={data.invoices.filter(i=>i.status==="Fulfilled"&&weekdayNames[invoiceDay(i).getDay()]===x.day)}/>} columns={[
                {key:"day",label:"Weekday"},{key:"orders",label:"Total orders",numeric:true},{key:"Dine In",label:"Dine In",numeric:true,render:x=>mixValue(x["Dine In"])},{key:"Pickup",label:"Pickup",numeric:true,render:x=>mixValue(x.Pickup)},{key:"Delivery",label:"Delivery",numeric:true,render:x=>mixValue(x.Delivery)},{key:"Talabat",label:"Talabat",numeric:true,render:x=>mixValue(x.Talabat)},{key:"Keeta",label:"Keeta",numeric:true,render:x=>mixValue(x.Keeta)},{key:"total",label:mixLabel,numeric:true,render:x=>mixValue(x.total)},{key:"dominant",label:"Dominant order type",render:x=><ChannelBadge name={x.dominant}/>},{key:"weakest",label:"Lowest active",render:x=><ChannelBadge name={x.weakest}/>}
              ]} totals={{day:"TOTAL",orders:num(+s.fulfilledInvoices),"Dine In":mixValue(weekdayOrderMix.reduce((a,x)=>a+x["Dine In"],0)),Pickup:mixValue(weekdayOrderMix.reduce((a,x)=>a+x.Pickup,0)),Delivery:mixValue(weekdayOrderMix.reduce((a,x)=>a+x.Delivery,0)),Talabat:mixValue(weekdayOrderMix.reduce((a,x)=>a+x.Talabat,0)),Keeta:mixValue(weekdayOrderMix.reduce((a,x)=>a+x.Keeta,0)),total:mixValue(weekdayOrderMix.reduce((a,x)=>a+x.total,0)),dominant:"Talabat",weakest:"Keeta"}}/>
              <Info title="Decision use" tone="blue">Click a weekday row to expand its order-type and payment mix from highest to lowest. Use this matrix to plan channel availability, kitchen capacity and promotions; it measures sales contribution, not profitability.</Info>
            </div>
          </section>
          <section className="analyticsGrid">
            <div className="panel wide" id="daily-insights"><div className="panelHead"><div><h3>Daily revenue trend · June 2026</h3><p>Day numbers run from 1 to 30; hover a point for the exact revenue.</p></div><Badge tone="good">30 DAYS</Badge></div>
              <MetricLine rows={businessInsights.daily.map(x=>({...x,dayLabel:String(Number(x.date.slice(-2))),displayValue:biMetric(x)}))} valueKey="displayValue" labelKey="dayLabel"/>
              <DataGrid title="Daily control table" description="Red rows contain an order/payment channel mismatch" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="daily-performance" rows={businessInsights.daily.map(x=>({...x,displayValue:biMetric(x),hasPaymentMismatch:hasChannelMismatch(invoicesForDay(x.date))}))} rowClassName={x=>x.hasPaymentMismatch?"paymentMismatchRow":""} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={invoicesForDay(x.date)}/>} columns={[
                {key:"date",label:"June day",numeric:true,value:x=>Number(x.date.slice(-2)),render:x=>String(Number(x.date.slice(-2)))},{key:"day",label:"Day"},{key:"orders",label:"Orders",numeric:true},{key:"guests",label:"Guests",numeric:true,defaultVisible:false},{key:"displayValue",label:mixLabel,numeric:true,render:x=>mixValue(x.displayValue)},{key:"discount",label:"Discounts",numeric:true,render:x=>money(x.discount)},{key:"vat",label:"VAT",numeric:true,defaultVisible:false,render:x=>money(x.vat)},{key:"averageCheck",label:"Avg check",numeric:true,render:x=>money(x.averageCheck)},{key:"firstActivity",label:"First bill",defaultVisible:false},{key:"lastActivity",label:"Last bill",defaultVisible:false},{key:"observedWindowMinutes",label:"Observed window",numeric:true,defaultVisible:false,render:x=>minutesLabel(x.observedWindowMinutes)}
              ]} totals={{date:"TOTAL",orders:num(+s.fulfilledInvoices),guests:num(+s.totalGuests),displayValue:mixValue(biMetric({orders:+s.fulfilledInvoices,revenue:+s.grossSales,vat:+s.vat})),discount:money(10461.25),vat:money(+s.vat),averageCheck:money(+s.averageCheck)}}/>
            </div>
            <div className="panel wide compactWeekly" id="weekly-insights"><div className="panelHead"><div><h3>Week-by-week control</h3><p>Revenue, orders and discount pressure across the month · hover each point for detail.</p></div></div>
              <MetricLine rows={businessInsights.weekly.map(x=>({...x,displayValue:biMetric(x)}))} valueKey="displayValue" labelKey="week" color="#6d55a3"/>
              <DataGrid title="Week-by-week control table" description="Click a week to expand order type × payment type" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="weekly-performance" rows={businessInsights.weekly.map((x,i)=>({...x,weekNumber:`Week ${i+1}`,displayValue:biMetric(x),hasPaymentMismatch:hasChannelMismatch(invoicesForWeek(x.week))}))} rowClassName={x=>x.hasPaymentMismatch?"paymentMismatchRow":""} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={invoicesForWeek(x.week)}/>} columns={[{key:"weekNumber",label:"Week no."},{key:"week",label:"Date range"},{key:"days",label:"Days",numeric:true},{key:"orders",label:"Orders",numeric:true},{key:"displayValue",label:mixLabel,numeric:true,render:x=>mixValue(x.displayValue)},{key:"discount",label:"Discounts",numeric:true,render:x=>money(x.discount)},{key:"averageCheck",label:"Avg check",numeric:true,render:x=>money(x.averageCheck)}]} totals={{weekNumber:"TOTAL",week:"June",days:"30",orders:num(+s.fulfilledInvoices),displayValue:mixValue(biMetric({orders:+s.fulfilledInvoices,revenue:+s.grossSales,vat:+s.vat})),discount:money(10461.25),averageCheck:money(+s.averageCheck)}}/>
              <div className="operatingCards"><div><span>Longest observed billing window</span><b>{businessInsights.operatingHighlights.longestDay.date}</b><strong>{businessInsights.operatingHighlights.longestDay.firstActivity}–{businessInsights.operatingHighlights.longestDay.lastActivity}</strong></div><div><span>Shortest observed billing window</span><b>{businessInsights.operatingHighlights.shortestDay.date}</b><strong>{businessInsights.operatingHighlights.shortestDay.firstActivity}–{businessInsights.operatingHighlights.shortestDay.lastActivity}</strong></div><div><span>Highest revenue date</span><b>{businessInsights.operatingHighlights.bestRevenueDay.date}</b><strong>{money(businessInsights.operatingHighlights.bestRevenueDay.revenue)}</strong></div><div><span>Most orders / recorded guests</span><b>{businessInsights.operatingHighlights.bestOrderDay.date}</b><strong>{businessInsights.operatingHighlights.bestOrderDay.orders} / {businessInsights.operatingHighlights.bestGuestDay.guests}</strong></div></div>
              <Info title="Operating-hours limitation" tone="amber">First and last bill times show the observed transaction window, not staff attendance or official opening hours. Near-24-hour windows should trigger a business-date-cutoff review before labor decisions are made.</Info>
            </div>
            <div className="panel wide weekPartPanel"><div className="panelHead"><div><h3>Weekday × weekend comparison</h3><p>Monday–Friday compared with Saturday–Sunday, using mutually exclusive UAE operating-week groups.</p></div><Badge tone="good">30 DAYS</Badge></div>
              <div className="weekPartLayout"><div className="weekPartChart">{weekPartRows.map(x=><article key={x.name}><div><ChannelBadge name={x.name}/><b>{x.days}</b><small>{x.dayCount} calendar days</small></div><span><i style={{width:`${x.averageDaily/Math.max(...weekPartRows.map(y=>y.averageDaily))*100}%`}}/></span><strong>{money(x.averageDaily)}<small>average daily sales</small></strong></article>)}</div><DataGrid title="Week-part control" description="Click a row for order type × payment type" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="week-part-performance" rows={weekPartRows} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={data.invoices.filter(i=>i.status==="Fulfilled"&&(x.name==="Weekend"?[0,6].includes(invoiceDay(i).getDay()):![0,6].includes(invoiceDay(i).getDay())))}/>} columns={[{key:"name",label:"Period"},{key:"days",label:"Days"},{key:"dayCount",label:"Calendar days",numeric:true},{key:"orders",label:"Orders",numeric:true},{key:"displayValue",label:mixLabel,numeric:true,render:x=>mixValue(x.displayValue)},{key:"averageDaily",label:"Avg daily sales",numeric:true,render:x=>money(x.averageDaily)},{key:"averageCheck",label:"Avg check",numeric:true,render:x=>money(x.averageCheck)}]} totals={{name:"TOTAL",days:"June",dayCount:"30",orders:num(+s.fulfilledInvoices),displayValue:mixValue(weekPartRows.reduce((a,x)=>a+x.displayValue,0)),averageDaily:money(+s.grossSales/30),averageCheck:money(+s.averageCheck)}}/></div>
            </div>
            <div className="panel wide channelGroupPanel"><div className="panelHead"><div><h3>Own Orders × Aggregator Orders</h3><p>Restaurant-controlled channels compared with online food aggregators. Revenue is controlled; profitability awaits commission and food-cost inputs.</p></div><Badge tone="good">{money(+s.grossSales)}</Badge></div>
              <div className="channelGroupLayout"><div className="channelGroupCards">{channelGroupRows.map(x=><button key={x.name} onClick={()=>{setOrderType(x.name==="Own Orders"?"All":"Talabat");document.getElementById("overview-order-mix")?.scrollIntoView({behavior:"smooth"})}}><span>{x.name}</span><small>{x.description}</small><strong>{money(x.revenue)}</strong><i><em style={{width:`${x.mix}%`}}/></i><b>{x.mix.toFixed(1)}% of sales · {num(x.orders)} orders · charges {x.charges?money(x.charges):"—"}</b></button>)}</div><DataGrid title="Own × aggregator control" description="Click a row to expand order type × payment type" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="channel-group-performance" rows={channelGroupRows} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={data.invoices.filter(i=>i.status==="Fulfilled"&&x.types.includes(i.orderType))}/>} columns={[{key:"name",label:"Channel group"},{key:"orders",label:"Orders",numeric:true},{key:"guests",label:"Guests",numeric:true},{key:"displayValue",label:mixLabel,numeric:true,render:x=>mixValue(x.displayValue)},{key:"discount",label:"Discounts",numeric:true,render:x=>money(x.discount)},{key:"charges",label:"Charges incl. VAT",numeric:true,render:x=>money(x.charges)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"averageCheck",label:"Avg check",numeric:true,render:x=>money(x.averageCheck)},{key:"mix",label:"Sales mix",numeric:true,render:x=>`${x.mix.toFixed(1)}%`}]} totals={{name:"TOTAL",orders:num(+s.fulfilledInvoices),guests:num(+s.totalGuests),displayValue:mixValue(channelGroupRows.reduce((a,x)=>a+x.displayValue,0)),discount:money(10461.25),charges:money(+s.charges),vat:money(+s.vat),averageCheck:money(+s.averageCheck),mix:"100.0%"}}/></div>
            </div>
            <div className="panel wide mealPanel" id="meal-time-insights"><div className="panelHead"><div><h3>Meal Time-Based Revenue Comparison</h3><p>Midnight, breakfast, lunch, snacks and dinner using the invoice settlement time. Click a meal slot to expand order type × payment type.</p></div><Badge tone="good">{money(+s.grossSales)}</Badge></div>
              <div className="mealVisual">{mealRows.map((x,i)=><button key={x.name} onClick={()=>document.querySelector(`.dataGrid-meal-time-performance tbody tr:nth-child(${i+1})`)?.scrollIntoView({behavior:"smooth",block:"center"})} title={`${x.name}: ${money(x.revenue)} · ${x.orders} orders`}><span><i style={{height:`${Math.max(2,x.revenue/Math.max(...mealRows.map(y=>y.revenue))*100)}%`}}/></span><b>{x.name}</b><strong>{x.orders?money(x.revenue):"—"}</strong><small>{x.orders?`${num(x.orders)} orders · ${x.mix.toFixed(1)}%`:"No fulfilled orders"}</small></button>)}</div>
              <DataGrid title="Meal slot control" description="Revenue includes each invoice’s respective charges" toolbarExtra={<MetricSelector value={mixMetric} onChange={setMixMetric}/>} id="meal-time-performance" rows={mealRows} renderExpanded={x=><MixDrill metric={mixMetric} paymentExceptions={paymentRecon?.exceptions} invoices={x.invoices}/>} columns={[{key:"name",label:"Meal slot"},{key:"timeSlot",label:"Time slot"},{key:"orders",label:"Orders",numeric:true},{key:"displayValue",label:mixLabel,numeric:true,render:x=>mixValue(x.displayValue)},{key:"discount",label:"Discounts",numeric:true,render:x=>money(x.discount)},{key:"charges",label:"Charges incl. VAT",numeric:true,render:x=>money(x.charges)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"averageCheck",label:"Avg check",numeric:true,render:x=>money(x.averageCheck)},{key:"mix",label:"Sales mix",numeric:true,render:x=>`${x.mix.toFixed(1)}%`}]} totals={{name:"TOTAL",timeSlot:"24 hours",orders:num(+s.fulfilledInvoices),displayValue:mixValue(mealRows.reduce((a,x)=>a+x.displayValue,0)),discount:money(10461.25),charges:money(+s.charges),vat:money(+s.vat),averageCheck:money(+s.averageCheck),mix:"100.0%"}}/>
            </div>
          </section>
          <section className="threeCol protectionGrid" id="protection-insights">
            <div className="panel"><div className="panelHead"><div><h3>Revenue leakage radar</h3><p>Commercial concessions and operational reversals.</p></div><Badge tone="bad">WATCH</Badge></div>
            <div className="leakageRows"><button onClick={()=>setTab("discounts")}><span>Total reductions</span><b>{money(businessInsights.summary.unifiedDiscount)}</b><small>Discounts plus separately titled Complimentary</small></button><button onClick={()=>setTab("operations")}><span>Cancelled KOT list value</span><b>{money(kotAudit?.summary.cancelledListValue||0)}</b><small>{num(kotAudit?.summary.cancelledLines||0)} cancelled lines</small></button><button onClick={()=>setTab("operations")}><span>Edited KOT lines</span><b>{num(kotAudit?.summary.editedLines||0)}</b><small>Previous/new values not exported</small></button><button onClick={()=>setTab("payments")}><span>Unallocated payments</span><b>{money(+s.paymentComponentGap)}</b><small>Reporting allocation gap—not lost revenue</small></button></div>
            </div>
            <div className="panel itemRankPanel"><div className="panelHead"><div><h3>Top items by quantity</h3><p>Demand and kitchen-volume leaders.</p></div></div><label className="rankSearch"><span>⌕</span><input value={itemRankQuery} onChange={e=>setItemRankQuery(e.target.value)} placeholder="Search any available item…"/></label><div className="rankList">{[...searchedRankItems].sort((a,b)=>b.qty-a.qty).slice(0,itemRankQuery?20:10).map(x=><div className="rank" key={x.name}><div><b>{x.name}</b><span>{num(x.qty)} sold</span></div><div className="grow"><Bar value={x.qty} max={Math.max(1,...data.topItems.map(y=>y.qty))}/></div><strong>{money(x.net)}</strong></div>)}</div>{!searchedRankItems.length&&<p className="rankEmpty">No item matches this search.</p>}</div>
            <div className="panel itemRankPanel"><div className="panelHead"><div><h3>Top items by revenue</h3><p>Realized sales after allocated discounts.</p></div></div><label className="rankSearch"><span>⌕</span><input value={itemRankQuery} onChange={e=>setItemRankQuery(e.target.value)} placeholder="Search any available item…"/></label><div className="rankList">{[...searchedRankItems].sort((a,b)=>b.net-a.net).slice(0,itemRankQuery?20:10).map(x=><div className="rank" key={x.name}><div><b>{x.name}</b><span>{num(x.qty)} sold</span></div><div className="grow"><Bar value={x.net} max={Math.max(1,...data.topItems.map(y=>y.net))} color="#d28e25"/></div><strong>{money(x.net)}</strong></div>)}</div>{!searchedRankItems.length&&<p className="rankEmpty">No item matches this search.</p>}</div>
          </section>
          <section className="panel controllerActions" id="controller-recommendations"><div className="panelHead"><div><h3>Controller recommendations from June performance</h3><p>Prioritized actions based on revenue, speed, concessions and control quality.</p></div><Badge tone="warn">ACTION PLAN</Badge></div><div className="actionGrid">
            <article><span>01 · Protect peak capacity</span><h4>Friday leads average daily revenue</h4><p>Average Friday revenue is {money([...businessInsights.weekday].sort((a,b)=>b.averageDailyRevenue-a.averageDailyRevenue)[0]?.averageDailyRevenue||0)}. Protect stock, kitchen labor and aggregator availability around the 2 PM revenue peak.</p></article>
            <article><span>02 · Repair delivery speed</span><h4>Delivery averages {minutesLabel(businessInsights.turnaroundByType.find(x=>x.orderType==="Delivery")?.averageMinutes)}</h4><p>This is substantially slower than Pickup at {minutesLabel(businessInsights.turnaroundByType.find(x=>x.orderType==="Pickup")?.averageMinutes)}. Split kitchen preparation, dispatch waiting and driver travel timestamps before assigning responsibility.</p></article>
            <article><span>03 · Govern concessions</span><h4>Discount pressure is {(businessInsights.summary.unifiedDiscount/(+s.correctedGrossBeforeDiscount)*100).toFixed(1)}%</h4><p>Review discount title × reason × user weekly. Separate promotional investment, approved complimentary service recovery and unauthorized leakage.</p></article>
            <article><span>04 · Classify cancellations</span><h4>{num(kotAudit?.summary.cancelledLines||0)} KOT cancellations require cause codes</h4><p>Distinguish customer change, duplicate punch, stock-out, kitchen delay and system error. Eight current events have no reason and every cancellation needs an action timestamp and approver.</p></article>
            <article><span>05 · Fix guest analytics</span><h4>Recorded guests closely mirror order count</h4><p>This suggests cover counts are not consistently maintained. Average-per-person and client-volume decisions should remain provisional until dine-in covers and customer identities are captured accurately.</p></article>
            <article><span>06 · Do not claim profit yet</span><h4>Revenue performance is not profitability</h4><p>Add recipe cost, packaging, aggregator commission and labor allocation. Then calculate contribution margin by item, order type, hour and business day.</p></article>
          </div></section>
          <Info title="Profitability boundary" tone="red">The dashboard can rank revenue, quantity, discounts, cancellations and contribution-risk indicators. It cannot call a day or item “most profitable” until recipe quantities, ingredient costs, packaging, aggregator commissions and labor allocation are configured. For now, “best-performing” means highest controlled revenue—not accounting profit.</Info>
        </>}
        <Info title="What this means" tone="blue">The dashboard headline, order-type, tax, discount and staff totals are reliable. Category mix, item revenue, split-payment exports and inventory profit metrics need redesign before they should guide decisions.</Info>
      </>}

      {(tab==="bills"||tab==="items")&&<section className="workspace">
        <div className="filters"><label className="search"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder={tab==="bills"?"Search bill, order ID, cashier or payment…":"Search item, SKU or bill…"} /></label>
          <select value={orderType} onChange={e=>setOrderType(e.target.value)}><option>All</option><option>Dine In</option><option>Pickup</option><option>Delivery</option><option>Talabat</option><option>Keeta</option></select>
          {tab==="bills"&&<select value={dateFilter} onChange={e=>setDateFilter(e.target.value)}><option value="All">All June dates</option>{[...new Set(data.invoices.map(x=>x.date))].map(d=><option key={d}>{d}</option>)}</select>}
          {tab==="bills"&&<><select value={controlFilter} onChange={e=>setControlFilter(e.target.value)}><option value="All">All controls</option><option value="pass">Pass</option><option value="corrected">Corrected</option><option value="review">Review</option></select><select value={discountFilter} onChange={e=>setDiscountFilter(e.target.value)}><option value="All">All discount checks</option><option value="pass">Discount pass</option><option value="corrected">Discount corrected</option><option value="rounding">Discount rounding</option></select><select value={vatFilter} onChange={e=>setVatFilter(e.target.value)}><option value="All">All VAT checks</option><option value="exact">VAT exact</option><option value="line-rounding">VAT line rounding</option></select></>}
          {tab==="items"&&<select value={priceStatus} onChange={e=>setPriceStatus(e.target.value)}><option>All</option><option value="match">match</option><option value="mismatch">mismatch</option><option value="unmatched">unmatched</option></select>}
          <span className="resultCount">{num(tab==="bills"?filteredInvoices.length:filteredItems.length)} results</span>
        </div>
        {tab==="bills"?<div className="split">
          <DataGrid id="bill-ledger" rows={filteredInvoices} selectedKey={selectedBill} onRowClick={x=>setSelectedBill(x.billNo)} columns={[
            {key:"billNo",label:"Bill",render:x=><><b>#{x.billNo}</b><small>{x.id}</small></>},{key:"date",label:"Date"},{key:"orderType",label:"Order type"},{key:"guests",label:"Guests",numeric:true,defaultVisible:false},
            {key:"canonicalGrossBeforeDiscount",label:"Original gross",numeric:true,render:x=>money(x.canonicalGrossBeforeDiscount)},
            {key:"canonicalItemDiscount",label:"Item discount",numeric:true,render:x=>money(x.canonicalItemDiscount)},
            {key:"canonicalOrderDiscount",label:"Order discount",numeric:true,render:x=>money(x.canonicalOrderDiscount)},
            {key:"canonicalTotalDiscount",label:"Total discount",numeric:true,render:x=>money(x.canonicalTotalDiscount)},
            {key:"canonicalTaxableInclVat",label:"Net incl VAT",numeric:true,render:x=>money(x.canonicalTaxableInclVat)},
            {key:"canonicalNetExVat",label:"Net ex VAT",numeric:true,defaultVisible:false,render:x=>money(x.canonicalNetExVat)},
            {key:"canonicalVat",label:"VAT",numeric:true,render:x=>money(x.canonicalVat)},{key:"charges",label:"Charges",numeric:true,render:x=>money(x.charges)},{key:"total",label:"Revenue",numeric:true,render:x=><b>{money(x.total)}</b>},
            {key:"discountAssessment",label:"Discount check",defaultVisible:false,render:x=><Badge tone={x.discountAssessment==="corrected"?"warn":"good"}>{x.discountAssessment}</Badge>},
            {key:"vatAssessment",label:"VAT check",defaultVisible:false,render:x=><Badge tone={x.vatAssessment==="exact"?"good":"warn"}>{x.vatAssessment}</Badge>},
            {key:"controlStatus",label:"Control",render:x=><Badge tone={x.controlStatus==="pass"?"good":x.controlStatus==="corrected"?"warn":"bad"}>{x.controlStatus}</Badge>}
          ]} totals={{billNo:"TOTAL",guests:num(filteredInvoices.reduce((a,x)=>a+x.guests,0)),canonicalGrossBeforeDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalGrossBeforeDiscount,0)),canonicalItemDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalItemDiscount,0)),canonicalOrderDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalOrderDiscount,0)),canonicalTotalDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalTotalDiscount,0)),canonicalTaxableInclVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalTaxableInclVat,0)),canonicalNetExVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalNetExVat,0)),canonicalVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalVat,0)),charges:money(filteredInvoices.reduce((a,x)=>a+x.charges,0)),total:money(filteredInvoices.reduce((a,x)=>a+x.total,0))}}/>
          <div className="drawer">{selected&&<><div className="drawerHead"><div><span>INVOICE TRACE</span><h3>Bill #{selected.billNo}</h3><p>{selected.date} · {selected.time} · {selected.orderType}</p></div>{selected.hasAnomaly?<Badge tone="bad">FAIL</Badge>:<Badge tone="good">PASS</Badge>}</div>
            <div className="invoiceTime"><div><span>ORDER PLACED</span><b>{selectedTurnaround?.placedAt||"Not captured"}</b></div><i>→</i><div><span>SETTLED</span><b>{selectedTurnaround?.settledAt||"Not captured"}</b></div><strong><span>CUSTOMER / ORDER TURNAROUND</span>{minutesLabel(selectedTurnaround?.turnaroundMinutes)}</strong></div>
            <div className="bridge"><div><span>Subtotal</span><b>{money(selected.subtotal)}</b></div><i>−</i><div><span>Invoice discount</span><b>{money(selected.discount)}</b></div><i>+</i><div><span>Charges</span><b>{money(selected.charges)}</b></div><i>=</i><div className="total"><span>Gross total</span><b>{money(selected.total)}</b></div></div>
            <div className="controlBox bad"><span>ITEM ↔ INVOICE DISCOUNT CONTROL</span><div><b>Item lines</b><strong>{money(selected.itemLineDiscount)}</strong></div><div><b>Invoice header</b><strong>{money(selected.discount)}</strong></div><div><b>Variance</b><strong>{money(selected.discountVariance)}</strong></div></div>
            <div className="controlBox"><span>CORRECTED DISCOUNT & VAT BRIDGE</span><div><b>Gross Item Sales</b><strong>{money(selected.canonicalGrossBeforeDiscount)}</strong></div><div><b>Less: Item Discounts</b><strong>− {money(selected.canonicalItemDiscount)}</strong></div><div><b>Less: Order Discounts</b><strong>− {money(selected.canonicalOrderDiscount)}</strong></div><div><b>Net Item Sales incl. VAT</b><strong>{money(selected.canonicalTaxableInclVat)}</strong></div><div><b>VAT included (5/105)</b><strong>{money(selected.canonicalVat)}</strong></div><div><b>Net Item Sales excl. VAT</b><strong>{money(selected.canonicalNetExVat)}</strong></div></div>
            <h4>Item lines</h4>{selectedLines.map(x=><div className={`lineCard ${x.anomaly?"anomaly":""}`} key={x.rowId}><div><b>{x.name}</b><small>{x.itemId} · Qty {x.qty} · {x.category}</small></div><div><span>Price</span><b>{money(x.actualPrice)}</b></div><div><span>Discount</span><b>{money(x.lineDiscount)}</b></div><div><span>Net export</span><b>{money(x.netAmount)}</b></div>{x.anomaly&&<Badge tone="bad">Negative taxable line</Badge>}</div>)}
            {selected.billNo==="2990"&&<Info title="Correct treatment for #2990" tone="red">Fattoush original value ⃃ 13.00 is fully removed as an item discount. It is excluded from the 10% order-discount base. The eligible ⃃ 81.00 receives ⃃ 8.10 order discount: original gross ⃃ 94.00 − item discount ⃃ 13.00 − order discount ⃃ 8.10 = taxable total ⃃ 72.90, VAT ⃃ 3.47 and net ex VAT ⃃ 69.43.</Info>}
            <div className="ebillHead"><div><h4>3-inch eBill preview</h4><p>Local audit preview · official receipt linked for verification</p></div><a href={`https://backoffice.tmbill.com/ebill/${selected.id}`} target="_blank" rel="noreferrer">Open official receipt ↗</a></div>
            <div className="receiptFrame local">
              <div className="receiptLocal">
                <div className="receiptAccent"/>
                <div className="receiptBrand">MORBIDO</div>
                <h3>Morbido Express Restaurant</h3>
                <p>Eastern Rd - Al Nahyan - E19 02 - Abu Dhabi - UAE</p>
                <p>TAX Invoice</p><p>TRN: 104855802500003</p>
                <div className="receiptRule"/>
                <p>{selected.date} · {selected.time}</p><b className="orderPill">{selected.orderType}</b>
                <h2>Order No: {selected.billNo}</h2><small>Order ID: {selected.id}</small>
                <div className="receiptSection"><h5>Order Details</h5><div><span>Status</span><b>{selected.status}</b></div><div><span>Table</span><b>{selected.table||"—"}</b></div><div><span>User</span><b>{selected.user}</b></div><div><span>Payment</span><b>{selected.paymentMode}</b></div></div>
                <div className="receiptSection"><h5>Order Items</h5><div className="receiptItem head"><span>Item</span><span>Qty</span><span>Rate</span><span>Total</span></div>
                  {selectedLines.map(x=><div className="receiptItem" key={x.rowId}><span>{x.name}{x.lineDiscount>0&&<small>Discount {money(x.lineDiscount)}</small>}</span><span>{x.qty}</span><span>{money(x.actualPrice===0&&x.lineDiscount>0?x.lineDiscount/x.qty:x.actualPrice).replace("AED","")}</span><span>{money(Math.max(0,x.netAmount)).replace("AED","")}</span></div>)}
                </div>
                <div className="receiptSection summary"><div><span>Subtotal</span><b>{money(selected.subtotal)}</b></div><div><span>Total Discount</span><b>{money(selected.discount)}</b></div><div><span>Total Without Tax</span><b>{money(selected.total-selected.vat)}</b></div><div><span>Total Tax</span><b>{money(selected.vat)}</b></div><div className="grand"><span>Grand Total</span><b>{money(selected.total)}</b></div></div>
                <div className="receiptSection"><h5>Payment Info</h5><div><span>{selected.paymentMode}</span><b>{money(selected.total)}</b></div></div>
                <p className="thanks">Thank You! 🙏</p><small>Powered by TMBill · Local audit reconstruction</small>
              </div>
            </div>
          </>}</div>
        </div>:<DataGrid id="item-ledger" rows={filteredItems} columns={[
          {key:"name",label:"Item",render:x=><><b>{x.name}</b><small>Bill #{x.billNo} · {x.itemId}</small></>},{key:"orderType",label:"Order type"},{key:"qty",label:"Qty",numeric:true},
          {key:"actualPrice",label:"Transaction list price",numeric:true,render:x=>money(x.actualPrice)},{key:"expectedCurrentPrice",label:"Current menu price",numeric:true,render:x=>x.expectedCurrentPrice===null?"Not matched":money(x.expectedCurrentPrice)},
          {key:"monthlyAvgGrossUnit",label:"Monthly avg list",numeric:true,defaultVisible:false,render:x=>money(x.monthlyAvgGrossUnit)},{key:"orderTypeAvgGrossUnit",label:"Order-type avg list",numeric:true,defaultVisible:false,render:x=>money(x.orderTypeAvgGrossUnit)},
          {key:"canonicalItemDiscount",label:"Item discount",numeric:true,render:x=>money(x.canonicalItemDiscount)},{key:"canonicalOrderDiscount",label:"Order discount",numeric:true,render:x=>money(x.canonicalOrderDiscount)},{key:"effectiveDiscountRate",label:"Effective discount %",numeric:true,render:x=>`${x.effectiveDiscountRate.toFixed(2)}%`},
          {key:"canonicalNetInclVat",label:"Sold incl VAT",numeric:true,render:x=>money(x.canonicalNetInclVat)},{key:"canonicalNetExVat",label:"Sold ex VAT",numeric:true,render:x=>money(x.canonicalNetExVat)},{key:"canonicalVat",label:"VAT",numeric:true,render:x=>money(x.canonicalVat)},
          {key:"monthlyAvgSoldInclVat",label:"Monthly avg sold incl VAT",numeric:true,defaultVisible:false,render:x=>money(x.monthlyAvgSoldInclVat)},{key:"monthlyAvgSoldExVat",label:"Monthly avg sold ex VAT",numeric:true,defaultVisible:false,render:x=>money(x.monthlyAvgSoldExVat)},
          {key:"orderTypeAvgSoldInclVat",label:"Order-type avg sold incl VAT",numeric:true,defaultVisible:false,render:x=>money(x.orderTypeAvgSoldInclVat)},{key:"orderTypeAvgSoldExVat",label:"Order-type avg sold ex VAT",numeric:true,defaultVisible:false,render:x=>money(x.orderTypeAvgSoldExVat)},
          {key:"priceStatus",label:"Menu comparison",render:x=><Badge tone={x.anomaly?"bad":x.priceStatus==="match"?"good":x.priceStatus==="mismatch"?"warn":"neutral"}>{x.anomaly?"Data anomaly":x.priceStatus}</Badge>}
        ]} totals={{name:"TOTAL",qty:num(filteredItems.reduce((a,x)=>a+x.qty,0)),canonicalItemDiscount:money(filteredItems.reduce((a,x)=>a+x.canonicalItemDiscount,0)),canonicalOrderDiscount:money(filteredItems.reduce((a,x)=>a+x.canonicalOrderDiscount,0)),canonicalNetInclVat:money(filteredItems.reduce((a,x)=>a+x.canonicalNetInclVat,0)),canonicalNetExVat:money(filteredItems.reduce((a,x)=>a+x.canonicalNetExVat,0)),canonicalVat:money(filteredItems.reduce((a,x)=>a+x.canonicalVat,0))}}/>}
        {tab==="items"&&<Info title="Price comparison limitation" tone="amber">The menu file is a current snapshot. Without effective-from/effective-to dates, a mismatch is an audit exception—not proof that the June transaction was incorrectly priced. Production should join by immutable item ID and the price version effective at supply time.</Info>}
      </section>}

      {tab==="discounts"&&<>
        {businessInsights&&<><section className="kpiGrid compact"><Kpi label="Discounts" value={money(businessInsights.summary.invoiceDiscount)} sub="Order and promotional reductions · 587 fulfilled invoices"/><Kpi label="Complimentary" value={money(businessInsights.summary.itemComplimentary)} sub="Fattoush on #2990 · authorized reason retained" tone="warn"/><Kpi label="Total reductions" value={money(businessInsights.summary.unifiedDiscount)} sub="Discounts + Complimentary"/><Kpi label="Reduction evidence rows" value={num(businessInsights.discountDetails.length)} sub="Title, reason, user, bill and source linked"/></section>
        <Info title="Standard treatment of the ⃃ 13" tone="blue">The ⃃ 13 is classified only as <b>Complimentary</b>, not as a normal discount. It is the Fattoush on bill #2990 with its recorded reason, actor and timestamp. The separate 10% “Al Zaeem” order discount is ⃃ 8.10 and must exclude the complimentary line from its eligible base. Total reductions are therefore ⃃ 10,461.25 Discounts + ⃃ 13.00 Complimentary = ⃃ 10,474.25.</Info>
        <section className="panel"><div className="panelHead"><div><h3>Discount title × reason summary</h3><p>Linked from Sales Report discount components and the confirmed complimentary event.</p></div><Badge tone="good">LINKED</Badge></div>
          <DataGrid id="discount-summary-linked" rows={businessInsights.discountSummary} columns={[
            {key:"title",label:"Discount title"},{key:"reason",label:"Recorded reason"},{key:"level",label:"Level"},{key:"bills",label:"Bills",numeric:true},{key:"amount",label:"Discount amount",numeric:true,render:x=>money(x.amount)},{key:"users",label:"Users"},{key:"sources",label:"Evidence source"}
          ]} totals={{title:"UNIFIED TOTAL",bills:num(new Set(businessInsights.discountDetails.map(x=>x.billNo)).size),amount:money(businessInsights.summary.unifiedDiscount)}}/>
        </section>
        <div className="filters discountFilters"><label className="search"><span>⌕</span><input value={discountQuery} onChange={e=>setDiscountQuery(e.target.value)} placeholder="Search bill, order ID, title, reason, user or order type…"/></label><select value={discountLevel} onChange={e=>setDiscountLevel(e.target.value)}><option>All</option><option>Order / invoice</option><option>Item</option></select><span className="resultCount">{num(filteredDiscountDetails.length)} evidence rows</span></div>
        <DataGrid id="discount-evidence" rows={pagedDiscountDetails} onRowClick={x=>{setSelectedBill(x.billNo);setTab("bills")}} columns={[
          {key:"billNo",label:"Bill",render:x=><><b>#{x.billNo}</b><small>{x.orderId}</small></>},{key:"date",label:"Date"},{key:"settledAt",label:"Settled time",defaultVisible:false},{key:"orderType",label:"Order type"},{key:"title",label:"Discount title"},{key:"reason",label:"Reason"},{key:"level",label:"Level"},{key:"amount",label:"Amount",numeric:true,render:x=>money(x.amount)},{key:"user",label:"User"},{key:"source",label:"Linked source",defaultVisible:false},{key:"control",label:"Control",render:x=><Badge tone={x.control==="review"?"bad":x.control==="confirmed"?"warn":"good"}>{x.control}</Badge>}
        ]} totals={{billNo:`PAGE ${discountPage}/${discountPageCount}`,title:`${num(pagedDiscountDetails.length)} displayed`,amount:money(pagedDiscountDetails.reduce((a,x)=>a+x.amount,0))}}/>
        <div className="pager"><button disabled={discountPage===1} onClick={()=>setDiscountPage(p=>p-1)}>← Previous</button><span>Rows {(discountPage-1)*discountPageSize+1}–{Math.min(discountPage*discountPageSize,filteredDiscountDetails.length)} of {num(filteredDiscountDetails.length)}</span><button disabled={discountPage===discountPageCount} onClick={()=>setDiscountPage(p=>p+1)}>Next →</button></div></>}
        <section className="panel"><div className="panelHead"><div><h3>Required discount allocation waterfall</h3><p>The exact calculation every category and item report should use.</p></div></div>
          <div className="steps">{[
            ["1","Line gross","Quantity × transaction unit price"],
            ["2","Item discount","Subtract explicit line-level discount"],
            ["3","Eligible base","Exclude non-discountable lines and partition by tax rate"],
            ["4","Allocate order discount","Order discount × eligible line gross ÷ eligible invoice gross"],
            ["5","Absorb rounding residual","Last eligible line receives the ⃃ 0.01 residual"],
            ["6","Validate","Allocated item + order discounts must equal invoice discount exactly"]
          ].map(x=><div className="step" key={x[0]}><span>{x[0]}</span><div><b>{x[1]}</b><p>{x[2]}</p></div></div>)}</div>
        </section>
        <Info title="Why this matters" tone="blue">Order-level discounts cannot remain only on the bill header if users expect category, kitchen, product-group and sold-item totals to match sales. They must be allocated to eligible item lines using a visible, deterministic rule.</Info>
      </>}

      {tab==="operations"&&kotAudit&&<>
        <section className="hero slim"><div><Badge tone="good">Six new source reports linked</Badge><h2>Item, KOT, cancellation and complimentary audit</h2><p>Trace what was punched, by whom, at what time, and what later changed. Cancelled events are operational controls and are never added to fulfilled sales.</p></div><div className="heroScore"><b>{num(kotAudit.summary.kotRows)}</b><span>KOT item events</span><small>{num(kotAudit.summary.billsWithKot)} bills linked</small></div></section>
        <section className="kpiGrid compact">
          <Kpi label="Placed item lines" value={num(kotAudit.summary.statusCounts.Placed||0)} sub="Production instructions retained"/>
          <Kpi label="Cancelled KOT lines" value={num(kotAudit.summary.cancelledLines)} sub={`${num(kotAudit.summary.cancelledQuantity)} quantity · list value ${money(kotAudit.summary.cancelledListValue)}`} tone="bad"/>
          <Kpi label="Missing cancel reason" value={num(kotAudit.summary.cancelReasonMissing)} sub="Must require reason + approver" tone="bad"/>
          <Kpi label="Edited KOT lines" value={num(kotAudit.summary.editedLines)} sub="Previous/new values are not exported" tone="warn"/>
          <Kpi label="Complimentary lines" value={num(kotAudit.summary.complimentaryLines)} sub={`${money(kotAudit.summary.complimentaryValue)} confirmed · bill #2990`} tone="warn"/>
          <Kpi label="KOT users" value={num(Object.keys(kotAudit.summary.users).length)} sub={Object.entries(kotAudit.summary.users).map(([u,c])=>`${u}: ${c}`).join(" · ")}/>
        </section>
        <Info title="New confirmation for bill #2990" tone="amber">The Complimentary Items report independently records Fattoush, quantity 1, ⃃ 13.00, created by tarekbmr on 12 Jun 2026 at 08:41:17 with reason “nnnno tomatto”. The KOT report has the same item/date/user/reason on #2990, settled at zero, so this is an exact contextual match. However, the complimentary report uses a different numeric KOT reference; developers must export the immutable order ID, bill number, KOT ID and line ID directly.</Info>
        <div className="filters kotFilters"><label className="search"><span>⌕</span><input value={kotQuery} onChange={e=>setKotQuery(e.target.value)} placeholder="Search bill, order ID, KOT, item, table, user or reason…"/></label>
          <select value={kotStatus} onChange={e=>setKotStatus(e.target.value)}><option>All</option><option>Placed</option><option>Cancelled</option><option>Edited</option></select>
          <select value={kotUser} onChange={e=>setKotUser(e.target.value)}><option>All</option>{Object.keys(kotAudit.summary.users).map(x=><option key={x}>{x}</option>)}</select>
          <select value={kotDate} onChange={e=>setKotDate(e.target.value)}><option>All</option>{[...new Set(kotAudit.events.map(x=>x.date))].sort().map(x=><option key={x}>{x}</option>)}</select>
          <span className="resultCount">{num(filteredKotEvents.length)} events</span>
        </div>
        <section className="kotWorkspace">
          <div>
            <DataGrid id="kot-events" rows={pagedKotEvents} selectedKey={selectedKotBill} onRowClick={x=>setSelectedKotBill(x.billNo)} columns={[
              {key:"billNo",label:"Bill",render:x=><><b>#{x.billNo}</b><small>{x.orderId}</small></>},{key:"date",label:"Bill date"},{key:"billTime",label:"Closed",defaultVisible:false},{key:"punchTime",label:"KOT punched"},
              {key:"kotNo",label:"KOT no.",defaultVisible:false},{key:"item",label:"Item"},{key:"qty",label:"Punched qty",numeric:true},{key:"settledQty",label:"Settled qty",numeric:true,defaultVisible:false},
              {key:"unitPrice",label:"Unit price",numeric:true,render:x=>money(x.unitPrice)},{key:"listedValue",label:"List value",numeric:true,render:x=>money(x.listedValue)},{key:"settledValue",label:"Settled value",numeric:true,render:x=>money(x.settledValue)},
              {key:"status",label:"Status",render:x=><Badge tone={x.status==="Cancelled"?"bad":x.status==="Edited"?"warn":"good"}>{x.status}</Badge>},{key:"user",label:"User"},{key:"table",label:"Table"},{key:"reason",label:"Reason",render:x=>x.reason||<span className="zeroDash">—</span>}
            ]} totals={{billNo:`PAGE ${kotPage}/${kotPageCount}`,item:`${num(pagedKotEvents.length)} displayed`,qty:num(pagedKotEvents.reduce((a,x)=>a+x.qty,0)),listedValue:money(pagedKotEvents.reduce((a,x)=>a+x.listedValue,0))}}/>
            <div className="pager"><button disabled={kotPage===1} onClick={()=>setKotPage(p=>p-1)}>← Previous</button><span>Rows {(kotPage-1)*kotPageSize+1}–{Math.min(kotPage*kotPageSize,filteredKotEvents.length)} of {num(filteredKotEvents.length)}</span><button disabled={kotPage===kotPageCount} onClick={()=>setKotPage(p=>p+1)}>Next →</button></div>
          </div>
          <aside className="kotDrawer">{selectedKotSummary?<><div className="drawerHead"><div><span>INVOICE PRODUCTION TRACE</span><h3>Bill #{selectedKotSummary.billNo}</h3><p>{selectedKotSummary.date} · closed {selectedKotSummary.billTime}</p></div><Badge tone={selectedKotSummary.assessment==="review"?"bad":"good"}>{selectedKotSummary.assessment}</Badge></div>
            <div className="timelineMeta"><div><span>Order ID</span><b>{selectedKotSummary.orderId}</b></div><div><span>User(s)</span><b>{selectedKotSummary.users}</b></div><div><span>KOTs / lines</span><b>{selectedKotSummary.kotCount} / {selectedKotSummary.lineCount}</b></div><div><span>First → last punch</span><b>{selectedKotSummary.firstPunch.slice(11)} → {selectedKotSummary.lastPunch.slice(11)}</b></div></div>
            <div className="eventTimeline">{selectedKotEvents.map(x=><article key={x.id} className={x.status==="Cancelled"?"cancelled":x.status==="Edited"?"edited":""}><i/><div><span>{x.punchTime||x.billTime}</span><b>{x.item} × {x.qty}</b><small>KOT {x.kotNo} · {x.user} · {money(x.listedValue)}</small>{x.reason&&<em>{x.reason}</em>}</div><Badge tone={x.status==="Cancelled"?"bad":x.status==="Edited"?"warn":"good"}>{x.status}</Badge></article>)}</div>
            <a className="downloadBtn" href={`https://backoffice.tmbill.com/ebill/${selectedKotSummary.orderId}`} target="_blank" rel="noreferrer">Open official eBill ↗</a>
          </>:<div className="noSource"><b>Select a KOT row</b><p>The bill’s complete punch sequence will appear here.</p></div>}</aside>
        </section>
        <section className="twoCol operationalPanels">
          <div className="panel"><div className="panelHead"><div><h3>Cancelled items requiring oversight</h3><p>All 103 events · red means removed from production/sale.</p></div><Badge tone="bad">{num(kotAudit.summary.cancelledLines)}</Badge></div>
            <DataGrid id="cancel-events" rows={kotAudit.cancellations} onRowClick={x=>{setSelectedKotBill(x.billNo);document.querySelector(".kotWorkspace")?.scrollIntoView({behavior:"smooth"})}} columns={[
              {key:"billNo",label:"Bill"},{key:"punchTime",label:"Punched time"},{key:"item",label:"Item"},{key:"qty",label:"Qty",numeric:true},{key:"listedValue",label:"List value",numeric:true,render:x=>money(x.listedValue)},{key:"user",label:"User"},{key:"reason",label:"Reason",render:x=>x.reason||<Badge tone="bad">Missing</Badge>}
            ]} totals={{billNo:"TOTAL",qty:num(kotAudit.summary.cancelledQuantity),listedValue:money(kotAudit.summary.cancelledListValue)}}/>
          </div>
          <div className="panel"><div className="panelHead"><div><h3>Complimentary & zero-settled control</h3><p>Actor, timestamp, reason and contextual invoice link.</p></div><Badge tone="warn">REVIEW</Badge></div>
            {kotAudit.complimentary.map(x=><article className="compCard" key={x.id}><div><span>COMPLIMENTARY</span><h3>{x.item} × {x.qty}</h3><p>Bill #{x.billNo} · {x.orderId}</p></div><strong>{money(x.value)}</strong><dl><dt>Created</dt><dd>{x.createdAt}</dd><dt>User</dt><dd>{x.createdBy}</dd><dt>Reason</dt><dd>{x.reason}</dd><dt>Link</dt><dd>{x.linkAssessment}</dd></dl><button onClick={()=>{setSelectedKotBill(x.billNo);document.querySelector(".kotWorkspace")?.scrollIntoView({behavior:"smooth"})}}>View invoice timeline</button></article>)}
            <Info title="Control design" tone="red">Complimentary is a discount/authorization event, not a zero menu price. Preserve original line gross, set a separate complimentary amount, require reason and approver, and exclude the line from any subsequent order-level discount base.</Info>
          </div>
        </section>
        <section className="sourceAudit"><div className="panelHead"><div><h3>Audit of the six new reports</h3><p>What each export proves, where it fails, and the developer correction.</p></div></div>
          <DataGrid id="new-report-audit" rows={kotAudit.reportAudit} columns={[
            {key:"report",label:"Report"},{key:"grain",label:"Grain"},{key:"rows",label:"Rows",numeric:true},{key:"status",label:"Assessment",render:x=><Badge tone={x.status==="fail"?"bad":x.status==="warn"?"warn":"good"}>{x.status}</Badge>},
            {key:"finding",label:"What the file shows",render:x=><span className="wrapCell">{x.finding}</span>},{key:"correction",label:"Required correction",render:x=><span className="wrapCell">{x.correction}</span>}
          ]}/>
        </section>
        <Info title="Performance architecture" tone="blue">Only 75 KOT events are rendered per page. Search is deferred while typing, filtered rows are memoized, and the invoice drawer renders only the selected bill. This avoids the visible lag caused by mounting all 3,524 rows and their controls at once.</Info>
      </>}

      {tab==="payments"&&<>
        <section className="kpiGrid compact"><Kpi label="Invoice/payment summary" value={money(108172.75)} sub="Correct total"/><Kpi label="Tax report components" value={money(106990.85)} sub="Incomplete tender allocations" tone="bad"/><Kpi label="Unallocated amount" value={money(1181.9)} sub="Secondary split tenders omitted" tone="bad"/><Kpi label="Payment allocations" value="1,831" sub="15 more than invoices due to split tenders"/></section>
        <section className="twoCol"><div className="panel"><div className="panelHead"><div><h3>Correct payment summary</h3><p>Dashboard and printed report.</p></div><Badge tone="good">PASS</Badge></div>{paymentTotals.map(([name,value,count])=><div className="rank" key={name}><div><ChannelBadge name={name}/><span>{count} allocations</span></div><div className="grow"><Bar value={value} max={53321.15} color={channelHex[name]}/></div><strong>{money(value)}</strong></div>)}</div>
          <div className="panel"><div className="panelHead"><div><h3>Broken tax-payment export</h3><p>Payment component bridge.</p></div><Badge tone="bad">FAIL</Badge></div>
            {[["Card",52998.35],["Cash",9870.8],["Talabat",42601.1],["Keeta",1520.6]].map(([n,v])=><div className="ledgerRow" key={n as string}><ChannelBadge name={n as string}/><b>{money(v as number)}</b></div>)}
            <div className="ledgerTotal"><span>Component total</span><b>{money(106990.85)}</b></div><div className="ledgerGap"><span>Missing allocation</span><b>{money(1181.9)}</b></div>
          </div></section>
        <Info title="Correct data model" tone="blue">Create one payment-allocation row per invoice × tender. A split invoice must have multiple rows. Payment totals must come from allocation amounts; the invoice’s primary payment label is not sufficient.</Info>
      </>}

      {tab==="crosscheck"&&<>
        <section className="hero slim"><div><Badge tone="bad">TMBill versus canonical controls</Badge><h2>See the reported number, correct number and root cause together.</h2><p>Sales/order type is a supply dimension. Payment type is a settlement dimension. The ⃃ 1,181.90 was collected correctly on 15 split bills but omitted from the Tax Submission report’s secondary-tender columns.</p></div><div className="heroScore bad"><b>{money(1181.90)}</b><span>split tender omitted from tax report</span><small>Customer payments are fully reconciled</small></div></section>
        <section className="kpiGrid compact downloadableKpis">
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="Correct invoice sales" value={money(108172.75)} sub="Supply ledger · click to download Excel"/></a>
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="TMBill payment-tax export" value={money(106990.85)} sub="Incomplete allocations · download detail" tone="bad"/></a>
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="Confirmed discount correction" value={money(13)} sub="Bill #2990 · download control" tone="warn"/></a>
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="Unclassified charge gap" value={money(61)} sub="⃃ 581 vs ⃃ 520 · download detail" tone="bad"/></a>
        </section>
        <section className="panel"><div className="panelHead"><div><h3>Order type × payment type × VAT</h3><p>Correct matrix rebuilt from the bill-wise payment allocations. Click headers to sort or Columns to configure.</p></div></div>
          <DataGrid id="order-payment-matrix" rows={paymentMatrix.map(x=>({...x,variance:x.sales-x.exported,validation:(x.orderType==="Talabat"&&x.keeta>0)||(x.orderType==="Dine In"&&x.talabat>0)?"Channel mismatch":Math.abs(x.sales-x.exported)>.01?"Missing allocation":"Pass"}))} columns={[
            {key:"orderType",label:"Order type",channel:true},{key:"orders",label:"Orders",numeric:true},{key:"card",label:"Card",numeric:true,render:x=><TenderValue row={x} tender="card"/>},{key:"cash",label:"Cash",numeric:true,render:x=><TenderValue row={x} tender="cash"/>},{key:"talabat",label:"Talabat",numeric:true,render:x=><TenderValue row={x} tender="talabat"/>},{key:"deliveroo",label:"Deliveroo",numeric:true,render:x=><TenderValue row={x} tender="deliveroo"/>},{key:"keeta",label:"Keeta",numeric:true,render:x=><TenderValue row={x} tender="keeta"/>},{key:"charges",label:"Charges incl. VAT",numeric:true,render:x=>money(x.charges)},{key:"sales",label:"Correct sales",numeric:true,render:x=>money(x.sales)},{key:"vat",label:"Invoice VAT",numeric:true,render:x=>money(x.vat)},{key:"exported",label:"Tax report captured",numeric:true,render:x=>money(x.exported)},{key:"variance",label:"Split tender omitted",numeric:true,render:x=>x.variance?money(x.variance):"—"},{key:"validation",label:"Validation",render:x=><Badge tone={x.validation==="Pass"?"good":"bad"}>{x.validation}</Badge>}
          ]} totals={{orderType:"TOTAL",orders:"1,816",card:money(53321.15),cash:money(10729.90),talabat:money(42601.10),deliveroo:"—",keeta:money(1520.60),charges:money(581),sales:money(108172.75),vat:money(5124.92),exported:money(106990.85),variance:money(1181.90)}}/>
        </section>
        <Info title="What the reconciliation means" tone="red"><ul><li><b>⃃ 875.90</b> is the secondary Card/Cash portion of 11 Dine In split bills omitted by the Tax Submission report.</li><li><b>⃃ 306.00</b> is the secondary Card/Cash portion of 4 Pickup split bills omitted by the same report.</li><li>Total split tender omitted from that export: <b>⃃ 1,181.90</b>. It is not unpaid revenue.</li><li>The correct bill/payment ledger already contains these amounts and reconciles to <b>⃃ 108,172.75</b>.</li><li><b>⃃ 117.00</b> of Dine In tendered as Talabat and <b>⃃ 65.60</b> of Talabat tendered as Keeta remain genuine channel-mapping reviews.</li><li>Payment allocation must never determine output VAT filing; the invoice supply ledger is the VAT control.</li></ul></Info>
        <section className="resolutionPanel"><div className="panelHead"><div><h3>Exact resolution for ⃃ 1,181.90</h3><p>The error is in the Tax Submission report’s payment extraction—not in invoice sales.</p></div><a className="downloadBtn" href="/downloads/TMBill_Reconciliation_Control.xlsx" download>Download reconciliation Excel ↓</a></div>
          <ol><li><b>Fix the Tax Submission report query/service.</b> It currently captures only one component of Card/Cash split bills. Join all active payment-allocation records by Order ID.</li><li><b>Use the correct grain.</b> One row must represent one invoice × tender allocation, including allocation ID, amount, status, reference and timestamp.</li><li><b>Add an invoice control.</b> For every fulfilled invoice: tender allocations + wallet + due = invoice total, adjusted for valid refunds/credit notes.</li><li><b>Backfill the 15 split invoices.</b> Mark them “Split confirmed – correct in bill ledger / incomplete in Tax Submission report,” not unpaid or missing.</li><li><b>Keep the 3 channel mismatches separate.</b> They require mapping review but have no monetary shortfall.</li><li><b>Rebuild the reports.</b> The payment matrix must total ⃃ 108,172.75. VAT filing remains ⃃ 5,124.92 from the invoice VAT ledger, not from payments.</li></ol>
        </section>
        {paymentRecon&&<section className="panel"><div className="panelHead"><div><h3>Affected invoices requiring payment-report correction</h3><p>{paymentRecon.summary.exceptionInvoices} exceptions: 15 missing allocations plus 3 cross-channel validations.</p></div></div><DataGrid id="payment-exceptions" rows={paymentRecon.exceptions} onRowClick={x=>{setSelectedBill(x.billNo);setTab("bills")}} columns={[
          {key:"billNo",label:"Bill",render:x=><b>#{x.billNo}</b>},{key:"orderId",label:"Order ID"},{key:"date",label:"Date"},{key:"orderType",label:"Order type"},{key:"paymentMode",label:"Payment mode"},{key:"correctTotal",label:"Correct bill tender",numeric:true,render:x=>money(x.correctTotal)},{key:"tmbillTotal",label:"Tax report captured",numeric:true,render:x=>money(x.tmbillTotal)},{key:"missing",label:"Secondary split omitted",numeric:true,render:x=>x.missing?money(x.missing):"—"},{key:"delta_card",label:"Omitted Card",numeric:true,render:x=>x.delta_card?money(x.delta_card):"—"},{key:"delta_cash",label:"Omitted Cash",numeric:true,render:x=>x.delta_cash?money(x.delta_cash):"—"},{key:"issueType",label:"Issue type"},{key:"reconciliationStatus",label:"Resolution",render:x=><Badge tone={x.channelMismatch?"bad":"good"}>{x.reconciliationStatus}</Badge>}
        ]} totals={{billNo:"CONTROL",missing:money(paymentRecon.summary.missing),reconciliationStatus:`${paymentRecon.summary.splitPaymentsConfirmed} split confirmed · ${paymentRecon.summary.channelMismatches} mapping reviews`}}/></section>}
        <Info title="Why the charge gap is ⃃ 61.00" tone="amber"><ul><li>The invoice-level Sales Report records <b>⃃ 581.00</b> under Total Charges.</li><li>The DSR Month-wise and Bill-wise reports expose only <b>⃃ 520.00</b> as Delivery/Total Charges.</li><li>The unresolved difference is therefore <b>⃃ 61.00</b>.</li><li>The Sales Report contains twelve columns all named “Delivery Charges,” so the missing ⃃ 61 cannot be mapped to a named charge type.</li><li>Fix: replace those duplicate columns with a normalized invoice-charge table containing charge ID, name, net, VAT rate, VAT and gross. Then require charge rows to reconcile exactly to invoice Total Charges.</li></ul></Info>
        <section className="sourceAudit"><div className="panelHead"><div><h3>Every supplied workbook assessed</h3><p>Select a report for its confirmed issue, correction and original header scope.</p></div></div><div className="auditLayout"><div className="auditList">{sourceAudits.map(a=><button key={a.name} className={selectedSource===a.name?"active":""} onClick={()=>setSelectedSource(a.name)}><Badge tone={a.status==="pass"?"good":a.status==="warn"?"warn":"bad"}>{a.status}</Badge><span>{a.name}</span></button>)}</div>{(()=>{const a=sourceAudits.find(x=>x.name===selectedSource)!;return <article className="auditDetail"><span>{a.file}</span><h2>{a.name}</h2><Badge tone={a.status==="pass"?"good":a.status==="warn"?"warn":"bad"}>{a.status.toUpperCase()}</Badge><h4>What TMBill gets wrong or omits</h4><p>{a.issue}</p><h4>Required correction</h4><p>{a.fix}</p><h4>Headers reviewed</h4><div className="headerCloud">{a.headers.split(", ").map(h=><Badge key={h}>{h}</Badge>)}</div></article>})()}</div></section>
      </>}

      {tab==="separate"&&separateAudit&&<>
        <section className="uploadAudit"><div><Badge tone="good">REUSABLE SALES AUDITOR</Badge><h2>Audit any TMBill Sales Report</h2><p>Upload a compatible Sales Report Excel file. It is analysed locally in your browser and never mixed with the June control dataset.</p></div><label className="uploadDrop"><input type="file" accept=".xlsx,.xls" onChange={e=>handleSalesUpload(e.target.files?.[0])}/><b>Choose Sales Report Excel</b><span>Required: Order ID, Bill, Subtotal, Discount, Tax, Charges, Total and Status.</span></label>{uploadError&&<div className="uploadError">{uploadError}</div>}</section>
        <section className="hero slim isolatedHero"><div><Badge tone="warn">ISOLATED DATASET · NOT COMBINED WITH JUNE</Badge><h2>Sales Report audit: {separateAudit.summary.periodFrom} – {separateAudit.summary.periodTo}</h2><p>Source: {separateAudit.summary.file}. Uploading another workbook replaces only this tab’s working dataset.</p></div><div className="heroScore"><b>{money(separateAudit.summary.total)}</b><span>fulfilled invoice revenue</span><small>{separateAudit.summary.fulfilled} bills · #{separateAudit.summary.billFrom}–#{separateAudit.summary.billTo}</small></div></section>
        <section className="kpiGrid">
          <Kpi label="Subtotal before discount" value={money(separateAudit.summary.subtotal)} sub="Report invoice subtotal · reconciled"/>
          <Kpi label="Discounts" value={money(separateAudit.summary.discount)} sub="Item and order reductions; complimentary is disclosed separately when identified"/>
          <Kpi label="Reported VAT" value={money(separateAudit.summary.vat)} sub={`Invoice aggregate control ${money(separateAudit.summary.vatControl)}`}/>
          <Kpi label="Charges" value={money(separateAudit.summary.charges)} sub="13 charged invoices · report footer shows only ⃃ 75" tone="bad"/>
          <Kpi label="Actual Sales excl. VAT" value={money(separateAudit.summary.netExVat)} sub="Actual Sales incl. VAT less reported VAT"/>
          <Kpi label="Grand total" value={money(separateAudit.summary.total)} sub="Subtotal − discounts + charges · exact"/>
          <Kpi label="Bills requiring review" value={num(separateAudit.summary.reviewBills)} sub="Charge tax, charge mapping, split detail or VAT control" tone="warn"/>
          <Kpi label="Potential bill sequence gaps" value={num(Math.max(0,separateAudit.summary.billTo-separateAudit.summary.billFrom+1-separateAudit.summary.rows))} sub="Require void/cancellation series evidence" tone="warn"/>
        </section>
        <section className="twoCol">
          <div className="panel"><div className="panelHead"><div><h3>What is correct</h3><p>Controls supported by the uploaded workbook.</p></div><Badge tone="good">PASS</Badge></div><ul className="findingList goodList"><li>{separateAudit.summary.rows} transaction rows; {separateAudit.summary.fulfilled} are Fulfilled.</li><li>Invoice arithmetic variance is {money(separateAudit.summary.totalControlVariance)}.</li><li>{money(separateAudit.summary.subtotal)} − {money(separateAudit.summary.discount)} + {money(separateAudit.summary.charges)} = {money(separateAudit.summary.total)}.</li><li>Reported VAT {money(separateAudit.summary.vat)} versus invoice control {money(separateAudit.summary.vatControl)}.</li><li>Order-type, payment-label and VAT summaries are rebuilt directly from invoices.</li></ul></div>
          <div className="panel"><div className="panelHead"><div><h3>What is incorrect or unsafe</h3><p>Confirmed report defects and missing audit evidence.</p></div><Badge tone="bad">REVIEW</Badge></div><ul className="findingList badList"><li>Repeated “Delivery Charges” headers do not identify charge types.</li><li>Visible charge components differ from Total Charges by {money(separateAudit.summary.chargeComponentGap)}.</li><li>Tax On Charges is {money(separateAudit.summary.taxOnCharges)} across {separateAudit.summary.chargeBills} charged invoices.</li><li>{separateAudit.summary.splitUnverifiable} split-payment invoices lack tender allocations.</li><li>{separateAudit.summary.channelMappingReviews} order/payment mappings require review.</li><li>{Math.max(0,separateAudit.summary.billTo-separateAudit.summary.billFrom+1-separateAudit.summary.rows)} potential bill numbers require void/cancel evidence.</li></ul></div>
        </section>
        <Info title="Core report verdict" tone="red">The invoice Total column passes only when its arithmetic variance is zero. Charge breakdown, charge VAT, split-payment detail, channel mappings and bill-series completeness remain separate controls; review the filtered invoices before filing or settlement.</Info>
        <section className="threeCol auditSummaries"><div className="panel"><div className="panelHead"><div><h3>Order types</h3></div></div>{separateAudit.orderTypes.map(x=><div className="ledgerRow" key={x.name}><span>{x.name} · {x.bills}</span><b>{money(x.total)}</b></div>)}</div><div className="panel"><div className="panelHead"><div><h3>Payment labels</h3></div></div>{separateAudit.payments.map(x=><div className="ledgerRow" key={x.name}><span>{x.name} · {x.bills}</span><b>{money(x.total)}</b></div>)}</div><div className="panel"><div className="panelHead"><div><h3>VAT and charges</h3></div></div><div className="ledgerRow"><span>Reported VAT</span><b>{money(separateAudit.summary.vat)}</b></div><div className="ledgerRow"><span>Aggregate VAT control</span><b>{money(separateAudit.summary.vatControl)}</b></div><div className="ledgerRow"><span>VAT variance</span><b>{money(separateAudit.summary.vatControlVariance)}</b></div><div className="ledgerRow"><span>Total Charges</span><b>{money(separateAudit.summary.charges)}</b></div><div className="ledgerRow"><span>Tax On Charges</span><b>{money(separateAudit.summary.taxOnCharges)}</b></div></div></section>
        <section className="panel"><div className="panelHead"><div><h3>Order type × payment type × VAT</h3><p>Unexpected aggregator mappings are highlighted red. Card,Cash split labels are excluded from tender columns until their component amounts are supplied.</p></div></div><DataGrid id="separate-order-payment" rows={separateMatrix} columns={[
          {key:"orderType",label:"Order type",channel:true},{key:"orders",label:"Orders",numeric:true},
          {key:"card",label:"Card",numeric:true,render:x=><TenderValue row={x} tender="card"/>},{key:"cash",label:"Cash",numeric:true,render:x=><TenderValue row={x} tender="cash"/>},
          {key:"talabat",label:"Talabat",numeric:true,render:x=><TenderValue row={x} tender="talabat"/>},{key:"deliveroo",label:"Deliveroo",numeric:true,render:x=><TenderValue row={x} tender="deliveroo"/>},{key:"keeta",label:"Keeta",numeric:true,render:x=><TenderValue row={x} tender="keeta"/>},
          {key:"correctSales",label:"Correct sales",numeric:true,render:x=>money(x.correctSales)},{key:"invoiceVat",label:"Invoice VAT",numeric:true,render:x=>money(x.invoiceVat)},{key:"taxReportCaptured",label:"Tax report captured",numeric:true,render:x=>money(x.taxReportCaptured)}
        ]} totals={{orderType:"TOTAL",orders:num(separateAudit.summary.fulfilled),card:money(separateMatrix.reduce((a,x)=>a+x.card,0)),cash:money(separateMatrix.reduce((a,x)=>a+x.cash,0)),talabat:money(separateMatrix.reduce((a,x)=>a+x.talabat,0)),deliveroo:money(separateMatrix.reduce((a,x)=>a+x.deliveroo,0)),keeta:money(separateMatrix.reduce((a,x)=>a+x.keeta,0)),correctSales:money(separateAudit.summary.total),invoiceVat:money(separateAudit.summary.vat),taxReportCaptured:money(separateMatrix.reduce((a,x)=>a+x.taxReportCaptured,0))}}/></section>
        <div className="filters"><label className="search"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search bill, order ID or order type…"/></label><select value={separateStatus} onChange={e=>setSeparateStatus(e.target.value)}><option value="All">All assessments</option><option value="pass">Pass</option><option value="review">Review</option></select><select value={separateIssue} onChange={e=>setSeparateIssue(e.target.value)}><option value="All">All issues</option><option>Charges present with zero Tax On Charges</option><option>Charge components do not equal Total Charges</option><option>Split payment amounts unavailable</option><option>VAT differs from aggregate 5/105 control</option></select><button className="downloadBtn" onClick={downloadSeparateAudit}>Download current audit ↓</button><span className="resultCount">{num(separateRows.length)} bills</span></div>
        <DataGrid id="separate-invoices" rows={separateRows} selectedKey={selectedSeparateBill} onRowClick={x=>{setSelectedSeparateBill(x.billNo);setTimeout(()=>document.getElementById("separate-bill-preview")?.scrollIntoView({behavior:"smooth",block:"start"}),60)}} columns={[
          {key:"billNo",label:"Bill",render:x=><b>#{x.billNo}</b>},{key:"id",label:"Order ID"},{key:"date",label:"Date"},{key:"orderType",label:"Order type"},{key:"paymentMode",label:"Payment label"},{key:"subtotal",label:"Subtotal",numeric:true,render:x=>money(x.subtotal)},{key:"discount",label:"Discount",numeric:true,render:x=>money(x.discount)},{key:"vat",label:"Reported VAT",numeric:true,render:x=>money(x.vat)},{key:"vatControl",label:"VAT control",numeric:true,render:x=>money(x.vatControl)},{key:"vatVariance",label:"VAT variance",numeric:true,render:x=>x.vatVariance?money(x.vatVariance):"—"},{key:"charges",label:"Total charges",numeric:true,render:x=>x.charges?money(x.charges):"—"},{key:"chargeParts",label:"Visible charge component",numeric:true,render:x=>x.chargeParts?money(x.chargeParts):"—"},{key:"taxOnCharges",label:"Tax on charges",numeric:true,render:x=>x.taxOnCharges?money(x.taxOnCharges):"—"},{key:"total",label:"Total",numeric:true,render:x=>money(x.total)},{key:"issues",label:"Issues",render:x=>x.issues.length?<span className="issueText">{x.issues.join(" · ")}</span>:"—"},{key:"assessment",label:"Assessment",render:x=><Badge tone={x.assessment==="pass"?"good":"warn"}>{x.assessment}</Badge>}
        ]} totals={{billNo:"TOTAL",subtotal:money(separateRows.reduce((a,x)=>a+x.subtotal,0)),discount:money(separateRows.reduce((a,x)=>a+x.discount,0)),vat:money(separateRows.reduce((a,x)=>a+x.vat,0)),charges:money(separateRows.reduce((a,x)=>a+x.charges,0)),total:money(separateRows.reduce((a,x)=>a+x.total,0))}}/>
        {separateSelected&&<section className="separatePreview" id="separate-bill-preview"><div className="previewTop"><div><span>INVOICE PREVIEW</span><h2>Bill #{separateSelected.billNo}</h2><p>{separateSelected.date} · {separateSelected.time} · <ChannelBadge name={separateSelected.orderType}/></p></div><div className="previewLinks"><a href={`https://backoffice.tmbill.com/ebill/${separateSelected.id}`} target="_blank" rel="noreferrer">Open official TMBill eBill ↗</a><button onClick={()=>setSelectedSeparateBill("")}>Close</button></div></div><div className="receiptFrame local"><div className="receiptLocal"><div className="receiptAccent"/><div className="receiptBrand">MORBIDO</div><h3>Morbido Express Restaurant</h3><p>TAX Invoice · Audit preview</p><div className="receiptRule"/><ChannelBadge name={separateSelected.orderType}/><h2>Order No: {separateSelected.billNo}</h2><small>Order ID: {separateSelected.id}</small><div className="receiptSection"><h5>Order Details</h5><div><span>Status</span><b>{separateSelected.status}</b></div><div><span>Table</span><b>{separateSelected.table||"—"}</b></div><div><span>User</span><b>{separateSelected.user}</b></div><div><span>Payment label</span><ChannelBadge name={separateSelected.paymentMode}/></div><div><span>Guests</span><b>{separateSelected.guests||"—"}</b></div></div><div className="receiptSection summary"><div><span>Subtotal</span><b>{money(separateSelected.subtotal)}</b></div><div><span>Discount</span><b>{money(separateSelected.discount)}</b></div><div><span>Charges</span><b>{money(separateSelected.charges)}</b></div><div><span>Net without VAT</span><b>{money(separateSelected.total-separateSelected.vat)}</b></div><div><span>VAT</span><b>{money(separateSelected.vat)}</b></div><div className="grand"><span>Grand Total</span><b>{money(separateSelected.total)}</b></div></div><p className="thanks">Audit reconstruction</p><small>Item lines and split-tender amounts are not present in this Sales Report.</small></div></div><Info title="Invoice assessment" tone={separateSelected.issues.length?"amber":"blue"}>{separateSelected.issues.length?separateSelected.issues.join(" · "):"This invoice passes the controls available in the uploaded Sales Report."}</Info></section>}
      </>}

      {tab==="zreport"&&<>
        <section className="receiptPreviewHero"><div><Badge tone="good">80 MM / 3 INCH MANAGEMENT REFERENCE</Badge><h2>June advanced Z-report preview</h2><p>A print-ready management close containing the controlled sales bridge, order-type economics, tenders, discounts, cancellations, turnaround, best-performing periods and POS operational flags.</p></div><button onClick={()=>window.print()}>Print 80 mm preview</button></section>
        <section className="zPreviewWorkspace">
          <div className="receipt80" id="management-z-receipt">
            <header><div className="receiptLogoCrop"><img src="/tmbill-logo-source.png" alt="TMBill"/></div><h2>MORBIDO EXPRESS RESTAURANT</h2><p>THE SS · MANAGEMENT Z REPORT</p><p>01 JUN 2026 12:45 PM – 30 JUN 2026 11:58 PM</p><b>CLOSED · RECONCILED</b></header>
            <div className="receiptRule">================================</div>
            <section><h3>SALES CONTROL</h3>{[
              ["Gross Item Sales",printMoney(118066)],["Discounts",`(${printMoney(10461.25)})`],["Complimentary",`(${printMoney(13)})`],["Total reductions",`(${printMoney(10474.25)})`],["Net Item Sales incl. VAT",printMoney(107591.75)],["VAT included",printMoney(5124.92)],["Net Item Sales excl. VAT",printMoney(102466.83)],["Charges incl. VAT",printMoney(581)]
            ].map(([a,b])=><div key={a}><span>{a}</span><b>{b}</b></div>)}<div className="receiptGrand"><span>ACTUAL SALES incl. VAT</span><b>{printMoney(108172.75)}</b></div></section>
            <section><h3>BILL & CUSTOMER CONTROL</h3><div><span>Bill series</span><b>#2313–#4141</b></div><div><span>Fulfilled bills</span><b>1,816</b></div><div><span>Cancelled bills</span><b>13 · {printMoney(1094.50)}</b></div><div><span>Recorded guests</span><b>1,817</b></div><div><span>Items sold</span><b>{num(+s.totalItemQuantity)}</b></div><div><span>Average cheque</span><b>{printMoney(+s.averageCheck)}</b></div><div><span>Average per person</span><b>{printMoney(+s.averagePerPerson)}</b></div><div><span>Avg items / bill</span><b>{(+s.totalItemQuantity/+s.fulfilledInvoices).toFixed(2)}</b></div></section>
            <section><h3>ORDER TYPE SUMMARY</h3>{zOrderRows.map(x=><article key={x.name} className={x.flag?"receiptFlagged":""}><h4>{x.name} · {num(x.orders)} orders{x.complimentary?` · Complimentary ${printMoney(x.complimentary)}`:""}{x.flag&&<em>POS FLAG</em>}</h4><div><span>Discounts</span><b>{x.discount?printMoney(x.discount):"—"}</b></div><div><span>Charges</span><b>{x.charges?printMoney(x.charges):"—"}</b></div><div><span>Net sales incl. VAT</span><b>{printMoney(x.revenue)}</b></div>{x.flag&&<p>Operational flag: order type and tender mapping require TMBill POS review.</p>}</article>)}<div className="receiptGrand"><span>TOTAL · Discounts / Charges / Net</span><b>{printMoney(10461.25)} / {printMoney(581)} / {printMoney(108172.75)}</b></div></section>
            <section><h3>PAYMENT SUMMARY</h3>{paymentTotals.map(([n,v,c])=><div key={n}><span>{n} · {num(c)} allocations</span><b>{printMoney(v)}</b></div>)}<div className="receiptGrand"><span>Total collected</span><b>{printMoney(108172.75)}</b></div><p className="receiptNote">Split tender allocations: 1,831 payment rows across 1,816 invoices.</p></section>
            <section><h3>CASH AUDIT SUMMARY</h3><div><span>Cash sales collected</span><b>{printMoney(10729.90)}</b></div><div><span>Expenses recorded</span><b>—</b></div><div><span>Expected cash</span><b>{printMoney(10729.90)}</b></div><div><span>Counted denominations</span><b>NOT CAPTURED</b></div><div><span>Cash variance</span><b>NOT ASSESSABLE</b></div></section>
            <section><h3>DISCOUNT & CANCELLATION</h3><div><span>Order/promotional discounts</span><b>{printMoney(10461.25)}</b></div><div><span>Complimentary</span><b>{printMoney(13)}</b></div><div><span>Cancelled KOT lines</span><b>{num(kotAudit?.summary.cancelledItems||103)}</b></div><div><span>Cancelled KOT list value</span><b>{printMoney(kotAudit?.summary.cancelledListValue||4794)}</b></div><div><span>Modified bills</span><b>18</b></div><div><span>Reprinted bills</span><b>402</b></div></section>
            <section><h3>MEAL TIME REVENUE</h3>{mealRows.map(x=><div key={x.name}><span>{x.name} · {x.timeSlot}</span><b>{num(x.orders)} · {printMoney(x.revenue)}</b></div>)}</section>
            <section><h3>PRODUCT / CATEGORY CONTROL</h3>{data.topItems.slice(0,5).map(x=><div key={x.name}><span>{x.name} · {num(x.qty)}</span><b>{printMoney(x.net)}</b></div>)}<div><span>Category/item charge allocation</span><b>SEPARATE · {printMoney(581)}</b></div></section>
            <section><h3>OPERATING PERFORMANCE</h3><div><span>Average / median turnaround</span><b>{minutesLabel(businessInsights?.summary.overallAverageMinutes)} / {minutesLabel(businessInsights?.summary.overallMedianMinutes)}</b></div><div><span>P90 turnaround</span><b>{minutesLabel(businessInsights?.summary.overallP90Minutes)}</b></div><div><span>Average daily revenue</span><b>{printMoney(+s.grossSales/30)}</b></div><div><span>Avg weekday / weekend</span><b>{printMoney(weekPartRows[0].averageDaily)} / {printMoney(weekPartRows[1].averageDaily)}</b></div><div><span>Best revenue date</span><b>Fri 12 Jun · {printMoney(5405.90)}</b></div><div><span>Peak revenue hour</span><b>02 PM · {printMoney(12518.10)}</b></div><div><span>Strongest week</span><b>22–28 Jun · {printMoney(28685.65)}</b></div></section>
            <section><h3>SUMMARY DATA COVERAGE</h3><div><span>Sales / bills / order types</span><b>AVAILABLE</b></div><div><span>Payments / discounts / items</span><b>AVAILABLE</b></div><div><span>Categories / KOT cancellations</span><b>AVAILABLE</b></div><div><span>Expenses / wallet / dues</span><b>NO SOURCE DATA</b></div><div><span>Delivery boy / waiter / kitchen</span><b>PARTIAL / MISSING</b></div><div><span>Denomination count / variance</span><b>NO SOURCE DATA</b></div></section>
            <section className="receiptWarnings"><h3>TMBILL POS OPERATIONAL FLAGS</h3><p>• Dine In tendered as Talabat: {printMoney(117)}</p><p>• Talabat orders tendered as Keeta: {printMoney(65.60)}</p><p>• Tax-payment export omitted secondary split tenders: {printMoney(1181.90)}</p><p>• Invoice charges {printMoney(581)} vs DSR-classified charges {printMoney(520)}: gap {printMoney(61)}</p><p>• Charge VAT classification requires confirmation.</p></section>
            <footer><div className="receiptRule">================================</div><b>MANAGEMENT SIGN-OFF</b><p>Sales ledger: PASS · VAT ledger: CONTROLLED</p><p>Payment export / charge mapping: REVIEW</p><div className="signLine">Manager ____________________</div><div className="signLine">Date / Time _________________</div><small>Generated by TMBill Audit Dashboard · Reference design</small></footer>
          </div>
          <aside className="receiptGuide"><span>WHAT MANAGEMENT RECEIVES</span><h2>One close report, four distinct controls</h2><div><b>01 · Supply</b><p>Gross, reductions, net, VAT, charges and actual revenue reconcile in one sequence.</p></div><div><b>02 · Operations</b><p>Orders, guests, cancellations, turnaround and period performance show how the restaurant operated.</p></div><div><b>03 · Collection</b><p>Payments are shown separately from sales and VAT, including the split-tender control.</p></div><div><b>04 · Exceptions</b><p>POS mapping and report-extraction defects are visibly flagged instead of silently buried.</p></div><Info title="Print boundary" tone="amber">This is a management Z-report reference. Tax filing should continue to use the governed invoice VAT ledger plus valid credit notes and documented adjustments.</Info></aside>
        </section>
        <section className="zSheet"><div className="zHead"><div className="zTitleBrand"><div className="zLogoCrop"><img src="/tmbill-logo-source.png" alt="TMBill Technology LLC"/></div><div><span>TAX / MANAGEMENT CONTROL REPORT</span><h2>Morbido Express Restaurant</h2><p>The SS · 01 Jun 2026 12:45 PM – 30 Jun 2026 11:58 PM</p></div></div><Badge tone="good">CLOSED · RECONCILED</Badge></div>
          <div className="zExecutiveStrip"><div><span>Actual sales incl. VAT</span><b>{printMoney(108172.75)}</b></div><div><span>VAT disclosed</span><b>{printMoney(5124.92)}</b></div><div><span>Net item sales excl. VAT</span><b>{printMoney(102466.83)}</b></div><div><span>Average cheque</span><b>{printMoney(+s.averageCheck)}</b></div></div>
          <div className="zMeta"><div><span>First / last bill</span><b>#2313 – #4141</b></div><div><span>Fulfilled bills</span><b>1,816</b></div><div><span>Cancelled bills</span><b>13 · {money(1094.50)}</b></div><div><span>Guests</span><b>1,817</b></div></div>
          <section className="zBlock"><h3>Sales and VAT bridge</h3>{[
            ["Gross Item Sales",118066],["Less: Discounts",-10461.25],["Less: Complimentary",-13],["Net Item Sales incl. VAT",107591.75],["VAT included",5124.92],["Net Item Sales excl. VAT",102466.83],["Charges incl. VAT",581],["ACTUAL SALES incl. VAT",108172.75]
          ].map(([n,v])=><div className={n==="ACTUAL SALES incl. VAT"?"grand":""} key={n as string}><span>{n}</span><b>{money(v as number)}</b></div>)}</section>
          <div className="zColumns"><section className="zBlock zOrderControl"><h3>Order type statistics</h3><div className="zOrderHeader"><span>Order type</span><b>Discounts</b><b>Charges</b><b>Net incl. VAT</b></div>{zOrderRows.map(x=><div key={x.name}><span><ChannelBadge name={x.name}/> ({x.orders}){x.complimentary?<small>Comp {money(x.complimentary)}</small>:null}</span><b>{x.discount?money(x.discount):"—"}</b><b>{x.charges?money(x.charges):"—"}</b><b>{money(x.revenue)}</b></div>)}<div className="grand"><span>TOTAL</span><b>{money(10461.25)}</b><b>{money(581)}</b><b>{money(108172.75)}</b></div></section>
          <section className="zBlock"><h3>Payment tenders</h3>{paymentTotals.map(([n,v,c])=><div key={n}><span>{n} ({c})</span><b>{money(v)}</b></div>)}<div className="grand"><span>Total</span><b>{money(108172.75)}</b></div></section></div>
          <section className="zBlock"><h3>Required review controls</h3><div><span>Confirmed item/order conflict</span><b>#2990 · corrected</b></div><div><span>Discount allocation rounding</span><b>6 invoices</b></div><div><span>VAT line rounding</span><b>252 invoices · max AED 0.02</b></div><div><span>Tax-payment allocation shortfall</span><b>{printMoney(1181.90)}</b></div><div><span>Charges without charge VAT</span><b>{printMoney(581)}</b></div></section>
          <Info title="Z-report sign-off" tone="amber">Do not use “payment collected” as the sales or VAT basis. Sign off the invoice supply ledger, discount bridge, VAT ledger, charge-tax classification and payment reconciliation separately. The AED 581 charge treatment remains unresolved until charge types/contracts are reviewed.</Info>
        </section>
      </>}

      {tab==="tax"&&<>
        <section className="hero slim"><div><Badge tone="good">No material invoice VAT exception found</Badge><h2>VAT total is plausible; the audit trail is not sufficient.</h2><p>Every closed invoice is tested against an inclusive 5% aggregate control. Differences are limited to AED ±0.02, consistent with line rounding. The item export still cannot reproduce the tax because its post-discount tax bases are unreliable.</p></div><div className="heroScore"><b>{money(5124.92)}</b><span>TMBill invoice VAT</span><small>Output VAT before credit-note adjustments</small></div></section>
        <section className="kpiGrid">
          <Kpi label="Taxable gross incl VAT" value={money(107591.75)} sub="Subtotal ⃃ 118,053 − discounts ⃃ 10,461.25"/>
          <Kpi label="TMBill reported VAT" value={money(5124.92)} sub="Sum of VAT on 1,816 supplied invoices"/>
          <Kpi label="Invoice aggregate control" value={money(5124.03)} sub="Σ round(invoice taxable gross × 5/105, 2)" tone="warn"/>
          <Kpi label="Single period control" value={money(5123.42)} sub="Round(period taxable gross × 5/105, 2)" tone="warn"/>
          <Kpi label="Exact invoice controls" value="1,564" sub="Reported VAT equals invoice aggregate control"/>
          <Kpi label="Rounding differences" value="252" sub="All differences are ⃃ 0.01 or ⃃ 0.02" tone="warn"/>
          <Kpi label="Material exceptions" value="0" sub="No invoice differs by more than ⃃ 0.02"/>
          <Kpi label="Invoice vs period rounding" value={money(1.50)} sub="Not automatically an under/overpayment" tone="warn"/>
        </section>
        <Info title="Critical compliance correction" tone="red">VAT payable is generally based on taxable supplies under the UAE date-of-supply rules—not only cash collected. An unpaid supplied invoice can still create output VAT. For this dataset, due receivables are zero, but the calculation engine must never equate “VAT payable” with “VAT collected in cash”.</Info>
        <Info title="Highest unresolved VAT risk" tone="amber">⃃ 581.00 of delivery/other charges is included in customer totals while “Tax on Charges” is ⃃ 0.00. If these charges are consideration for, or ancillary to, the standard-rated restaurant supply, they may also require VAT. If treated as VAT-inclusive, the potential VAT is ⃃ 27.67. TMBill must store each charge type, tax treatment, rate and VAT amount; a UAE tax adviser should confirm the classification before filing.</Info>
        <Info title="Two decimals or three?" tone="blue">Store calculations internally at high precision (at least 4–6 decimal places), but round the VAT amount payable on each tax invoice to the nearest fils—two decimal places—using mathematical half-up rounding. Three-decimal VAT can be retained only as an internal calculation trace; it should not be the posted invoice or filing amount. The system must document whether it rounds per line or at invoice total and use that method consistently.</Info>
        <Info title="Are service and delivery charges taxable?" tone="red">Generally, a compulsory service, delivery, packaging or similar fee imposed by the restaurant as part of making the taxable food supply forms part of consideration and follows the supply’s VAT treatment. A true disbursement paid in the customer’s name and account can differ. Because TMBill provides twelve anonymous “Delivery Charges” columns and zero Tax On Charges, the ⃃ 581 cannot be safely classified. The immediate fix is a charge ledger with charge name, compulsory/optional flag, principal/agent treatment, VAT rate, net, VAT and gross.</Info>
        <div className="filters"><label className="search"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search bill, order ID, cashier or payment…"/></label><select value={dateFilter} onChange={e=>setDateFilter(e.target.value)}><option value="All">All June dates</option>{[...new Set(data.invoices.map(x=>x.date))].map(d=><option key={d}>{d}</option>)}</select><select value={orderType} onChange={e=>setOrderType(e.target.value)}><option>All</option><option>Dine In</option><option>Pickup</option><option>Delivery</option><option>Talabat</option><option>Keeta</option></select><select value={vatFilter} onChange={e=>setVatFilter(e.target.value)}><option value="All">All VAT checks</option><option value="exact">Exact</option><option value="line-rounding">Line rounding</option></select><select value={discountFilter} onChange={e=>setDiscountFilter(e.target.value)}><option value="All">All discounts</option><option value="pass">Discount pass</option><option value="corrected">Corrected</option><option value="rounding">Discount rounding</option></select><span className="resultCount">{num(filteredInvoices.length)} invoices</span></div>
        <DataGrid id="vat-ledger" rows={filteredInvoices} onRowClick={x=>{setSelectedBill(x.billNo);setTab("bills")}} columns={[
          {key:"billNo",label:"Bill / order ID",render:x=><><b>#{x.billNo}</b><small>{x.id}</small></>},{key:"date",label:"Supply date"},{key:"orderType",label:"Order type"},
          {key:"canonicalGrossBeforeDiscount",label:"Original gross",numeric:true,render:x=>money(x.canonicalGrossBeforeDiscount)},{key:"canonicalItemDiscount",label:"Item discount",numeric:true,render:x=>money(x.canonicalItemDiscount)},{key:"canonicalOrderDiscount",label:"Order discount",numeric:true,render:x=>money(x.canonicalOrderDiscount)},
          {key:"canonicalTaxableInclVat",label:"Taxable incl VAT",numeric:true,render:x=>money(x.canonicalTaxableInclVat)},{key:"vat",label:"TMBill VAT",numeric:true,render:x=>money(x.vat)},{key:"canonicalVat",label:"Canonical line VAT",numeric:true,render:x=>money(x.canonicalVat)},{key:"taxControl",label:"Invoice aggregate VAT",numeric:true,render:x=>money(x.taxControl)},
          {key:"taxVariance",label:"TMBill vs aggregate",numeric:true,render:x=>money(x.taxVariance)},{key:"canonicalNetExVat",label:"Net taxable ex VAT",numeric:true,render:x=>money(x.canonicalNetExVat)},
          {key:"vatAssessment",label:"Assessment",render:x=><Badge tone={x.vatAssessment==="exact"?"good":"warn"}>{x.vatAssessment}</Badge>}
        ]} totals={{billNo:"TOTAL",canonicalGrossBeforeDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalGrossBeforeDiscount,0)),canonicalItemDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalItemDiscount,0)),canonicalOrderDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalOrderDiscount,0)),canonicalTaxableInclVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalTaxableInclVat,0)),vat:money(filteredInvoices.reduce((a,x)=>a+x.vat,0)),canonicalVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalVat,0)),taxControl:money(filteredInvoices.reduce((a,x)=>a+x.taxControl,0)),taxVariance:money(filteredInvoices.reduce((a,x)=>a+x.taxVariance,0)),canonicalNetExVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalNetExVat,0))}}/>
        <section className="twoCol taxExplain"><div className="panel"><div className="panelHead"><div><h3>What should be filed?</h3><p>Recommended output-tax basis.</p></div></div><div className="formulaList single"><div><code>Supplied invoice VAT ledger</code><b>{money(5124.92)}</b></div><div><code>Less valid credit notes</code><b>None supplied</b></div><div><code>Plus/minus documented adjustments</code><b>None supplied</b></div><div><code>Candidate output VAT</code><b>{money(5124.92)}</b></div></div></div>
          <div className="panel"><div className="panelHead"><div><h3>Core tax-reporting defect</h3><p>Why the current export remains risky.</p></div></div><p className="explainText">TMBill exposes invoice VAT but does not provide a dependable line-by-line bridge containing original line gross, item discount, allocated order discount, post-discount taxable amount, VAT rate and line VAT. Recalculating from the item export produces false exceptions because some order discounts and complimentary lines are not reflected consistently in its VAT-base fields.</p></div></section>
        <Info title="Required tax report columns" tone="blue">Invoice ID, bill number, supply date/time, status, taxable gross inclusive VAT, exempt/zero-rated gross, item discount, allocated order discount, taxable net excluding VAT, item VAT, charge VAT, total VAT, credit-note VAT, adjustment VAT, payment status and filing period. The invoice must drill into lines with the same fields.</Info>
      </>}

      {tab==="categories"&&<>
        <section className="hero slim"><div><Badge tone="good">Standardized category sales bridge</Badge><h2>Every category uses the same governed sales vocabulary.</h2><p>Gross Item Sales − Discounts − Complimentary = Net Item Sales incl. VAT. VAT is disclosed separately; Charges incl. VAT are then added to reach Actual Sales incl. VAT.</p></div><div className="heroScore"><b>{money(+s.taxableGrossInclVat)}</b><span>Net Item Sales incl. VAT</span><small>Charges incl. VAT: {money(+s.charges)}</small></div></section>
        <DataGrid id="category-ledger" rows={data.categories} columns={[
          {key:"name",label:"Category"},{key:"qty",label:"Qty",numeric:true},{key:"gross",label:"Gross Item Sales",numeric:true,render:x=>money(x.gross)},{key:"orderDiscount",label:"Discounts",numeric:true,render:x=>money(x.orderDiscount)},{key:"itemDiscount",label:"Complimentary",numeric:true,render:x=>money(x.itemDiscount)},{key:"net",label:"Net Item Sales incl. VAT",numeric:true,render:x=>money(x.net)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"netExVat",label:"Net Item Sales excl. VAT",numeric:true,render:x=>money(x.netExVat)},{key:"mix",label:"Mix %",numeric:true,value:x=>x.net/+s.taxableGrossInclVat*100,render:x=>`${(x.net/+s.taxableGrossInclVat*100).toFixed(2)}%`}
        ]} totals={{name:"TOTAL",qty:num(data.categories.reduce((a,c)=>a+c.qty,0)),gross:money(data.categories.reduce((a,c)=>a+c.gross,0)),itemDiscount:money(data.categories.reduce((a,c)=>a+c.itemDiscount,0)),orderDiscount:money(data.categories.reduce((a,c)=>a+c.orderDiscount,0)),discount:money(data.categories.reduce((a,c)=>a+c.discount,0)),net:money(data.categories.reduce((a,c)=>a+c.net,0)),vat:money(data.categories.reduce((a,c)=>a+c.vat,0)),netExVat:money(data.categories.reduce((a,c)=>a+c.netExVat,0)),mix:"100.00%"}}/>
        <Info title="Reconciliation rule" tone="blue">Net Item Sales incl. VAT reconcile to ⃃ 107,591.75. Add ⃃ 581.00 Charges incl. VAT to reach ⃃ 108,172.75 Actual Sales incl. VAT. Charges must not be forced into categories without an explicit allocation rule.</Info>
      </>}

      {tab==="reports"&&<>
        <section className="hero slim"><div><Badge tone="good">20 governed views</Badge><h2>Ideal restaurant reporting library</h2><p>Each view below uses the same invoice, item, payment and event facts. Select a report to understand its purpose and grain.</p></div></section>
        <Info title="Standard sales vocabulary" tone="blue"><b>Gross Item Sales</b> is item value before reductions. <b>Discounts</b> are order/promotional reductions. <b>Complimentary</b> is an authorized free item with title, reason, user and approval; the ⃃ 13 Fattoush on #2990 belongs only here. <b>Net Item Sales incl. VAT</b> equals gross less both reductions. <b>VAT</b> is the tax included in taxable sales and is disclosed, not deducted again from Actual Sales. <b>Net Item Sales excl. VAT</b> is the revenue base excluding VAT. <b>Charges incl. VAT</b> are separately classified invoice charges. <b>Actual Sales incl. VAT</b> equals Net Item Sales incl. VAT plus Charges incl. VAT and is the invoice/customer total.</Info>
        <div className="filters"><label className="search"><span>⌕</span><input value={reportSearch} onChange={e=>setReportSearch(e.target.value)} placeholder="Search reports and definitions…"/></label><span className="resultCount">{reportSpecs.filter(x=>x.join(" ").toLowerCase().includes(reportSearch.toLowerCase())).length} reports</span></div>
        <section className="reportGrid">{reportSpecs.filter(x=>x.join(" ").toLowerCase().includes(reportSearch.toLowerCase())).map((r,i)=><button className={`reportCard ${selectedReport===r[0]?"active":""}`} onClick={()=>{setSelectedReport(r[0]);setTimeout(()=>document.getElementById("ideal-report-preview")?.scrollIntoView({behavior:"smooth",block:"start"}),80)}} key={r[0]}><span>{String(i+1).padStart(2,"0")}</span><div><h3>{r[0]}</h3><Badge>{r[1]}</Badge><p>{r[2]}</p><strong>Preview report ↓</strong></div></button>)}</section>
        <section className="reportPreview" id="ideal-report-preview"><div className="previewTop"><div><span>REPORT PREVIEW</span><h2>{selectedReport}</h2><p>{reportSpecs.find(x=>x[0]===selectedReport)?.[2]}</p></div><div className="previewLinks"><a href="/downloads/TMBill_Standardized_Reporting_Control_Pack.xlsx" download>Download full standardized pack ↓</a><Badge tone="good">Canonical design</Badge></div></div>
          <div className="previewControls"><div><small>OUTLET</small><b>Morbido Express Restaurant</b></div><div><small>PERIOD</small><b>01–30 Jun 2026</b></div><div><small>STATUS</small><b>Supplied / fulfilled</b></div><div><small>BASIS</small><b>VAT inclusive + explicit VAT</b></div></div>
          {selectedReport==="Sales / Z Summary"?<DataGrid id="report-z" rows={[{period:"01–30 Jun 2026",bills:+s.fulfilledInvoices,gross:+s.correctedGrossBeforeDiscount,itemDiscount:+s.canonicalItemDiscount,orderDiscount:+s.canonicalOrderDiscount,charges:+s.charges,vat:+s.vat,netExVat:+s.netSales,total:+s.grossSales}]} columns={[
            {key:"period",label:"Period"},{key:"bills",label:"Bills",numeric:true},{key:"gross",label:"Gross Item Sales",numeric:true,render:x=>money(x.gross)},{key:"orderDiscount",label:"Discounts",numeric:true,render:x=>money(x.orderDiscount)},{key:"itemDiscount",label:"Complimentary",numeric:true,render:x=>money(x.itemDiscount)},{key:"netItems",label:"Net Item Sales incl. VAT",numeric:true,value:x=>x.gross-x.itemDiscount-x.orderDiscount,render:x=>money(x.gross-x.itemDiscount-x.orderDiscount)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"netExVat",label:"Net Item Sales excl. VAT",numeric:true,render:x=>money(x.netExVat-x.charges)},{key:"charges",label:"Charges incl. VAT",numeric:true,render:x=>money(x.charges)},{key:"total",label:"Actual Sales incl. VAT",numeric:true,render:x=>money(x.total)}
          ]}/>
          :selectedReport==="Payment Type Summary"?<DataGrid id="report-payments" rows={paymentTotals.map(([name,total,allocations])=>({name,total,allocations,mix:total/+s.grossSales*100}))} columns={[{key:"name",label:"Tender"},{key:"allocations",label:"Allocations",numeric:true},{key:"total",label:"Collected",numeric:true,render:x=>money(x.total)},{key:"mix",label:"Mix %",numeric:true,render:x=>`${x.mix.toFixed(2)}%`}]} totals={{name:"TOTAL",allocations:"1,831",total:money(+s.grossSales),mix:"100.00%"}}/>
          :selectedReport==="Discount Summary"&&businessInsights?<><DataGrid id="report-discounts" rows={businessInsights.discountSummary} columns={[{key:"title",label:"Discount title"},{key:"reason",label:"Recorded reason"},{key:"level",label:"Level"},{key:"bills",label:"Bills",numeric:true},{key:"amount",label:"Discount",numeric:true,render:x=>money(x.amount)},{key:"users",label:"Users"},{key:"sources",label:"Linked source"}]} totals={{title:"UNIFIED TOTAL",amount:money(businessInsights.summary.unifiedDiscount)}}/><Info title="Linked discount rule" tone="blue">The complimentary ⃃ 13 is confirmed and classified, not left as an unexplained variance. Production should preserve this item-level event separately from the ⃃ 8.10 Al Zaeem order discount on #2990 and exclude the complimentary line from the order-discount base.</Info></>
          :selectedReport==="Bill Summary"?<DataGrid id="report-bills" rows={data.invoices} columns={[
            {key:"billNo",label:"Bill"},{key:"id",label:"Order ID"},{key:"date",label:"Supply date"},{key:"orderType",label:"Order type"},
            {key:"canonicalGrossBeforeDiscount",label:"Gross Item Sales",numeric:true,render:x=>money(x.canonicalGrossBeforeDiscount)},
            {key:"canonicalOrderDiscount",label:"Discounts",numeric:true,render:x=>money(x.canonicalOrderDiscount)},
            {key:"canonicalItemDiscount",label:"Complimentary",numeric:true,render:x=>money(x.canonicalItemDiscount)},
            {key:"canonicalTaxableInclVat",label:"Net Item Sales incl. VAT",numeric:true,render:x=>money(x.canonicalTaxableInclVat)},
            {key:"canonicalVat",label:"VAT",numeric:true,render:x=>money(x.canonicalVat)},{key:"canonicalNetExVat",label:"Net Item Sales excl. VAT",numeric:true,render:x=>money(x.canonicalNetExVat)},{key:"charges",label:"Charges incl. VAT",numeric:true,render:x=>money(x.charges)},{key:"total",label:"Actual Sales incl. VAT",numeric:true,render:x=>money(x.total)}
          ]} totals={{billNo:"TOTAL",canonicalGrossBeforeDiscount:money(+s.correctedGrossBeforeDiscount),canonicalItemDiscount:money(+s.canonicalItemDiscount),canonicalOrderDiscount:money(+s.canonicalOrderDiscount),canonicalVat:money(5124.93),charges:money(+s.charges),total:money(+s.grossSales)}}/>
          :selectedReport==="Sold Items Summary"?<><DataGrid id="report-items" rows={data.topItems} columns={[
            {key:"name",label:"Item"},{key:"qty",label:"Qty",numeric:true},{key:"gross",label:"Gross Item Sales",numeric:true,render:x=>money(x.gross)},{key:"discount",label:"Total reductions",numeric:true,render:x=>money(x.discount)},{key:"net",label:"Net Item Sales incl. VAT",numeric:true,render:x=>money(x.net)}
          ]} totals={{name:"ITEM SALES TOTAL",qty:num(data.topItems.reduce((a,x)=>a+x.qty,0)),gross:money(data.topItems.reduce((a,x)=>a+x.gross,0)),discount:money(data.topItems.reduce((a,x)=>a+x.discount,0)),net:money(data.topItems.reduce((a,x)=>a+x.net,0))}}/><SalesReconciliation grossItems={data.topItems.reduce((a,x)=>a+x.gross,0)} totalReductions={data.topItems.reduce((a,x)=>a+x.discount,0)} complimentary={+(businessInsights?.summary.itemComplimentary||0)} vat={+s.vat} charges={+s.charges} actualSales={+s.grossSales}/></>
          :selectedReport==="Category Summary"||selectedReport==="Product Group Summary"?<><DataGrid id="report-category" rows={data.categories} columns={[
            {key:"name",label:"Category"},{key:"qty",label:"Qty",numeric:true},{key:"gross",label:"Gross Item Sales",numeric:true,render:x=>money(x.gross)},{key:"orderDiscount",label:"Discounts",numeric:true,render:x=>money(x.orderDiscount)},{key:"itemDiscount",label:"Complimentary",numeric:true,render:x=>money(x.itemDiscount)},{key:"net",label:"Net Item Sales incl. VAT",numeric:true,render:x=>money(x.net)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"netExVat",label:"Net Item Sales excl. VAT",numeric:true,render:x=>money(x.netExVat)}
          ]} totals={{name:"ITEM SALES TOTAL",qty:num(data.categories.reduce((a,x)=>a+x.qty,0)),gross:money(data.categories.reduce((a,x)=>a+x.gross,0)),orderDiscount:money(data.categories.reduce((a,x)=>a+x.orderDiscount,0)),itemDiscount:money(data.categories.reduce((a,x)=>a+x.itemDiscount,0)),net:money(data.categories.reduce((a,x)=>a+x.net,0)),vat:money(data.categories.reduce((a,x)=>a+x.vat,0)),netExVat:money(data.categories.reduce((a,x)=>a+x.netExVat,0))}}/><SalesReconciliation grossItems={data.categories.reduce((a,x)=>a+x.gross,0)} totalReductions={data.categories.reduce((a,x)=>a+x.discount,0)} complimentary={data.categories.reduce((a,x)=>a+x.itemDiscount,0)} vat={data.categories.reduce((a,x)=>a+x.vat,0)} charges={+s.charges} actualSales={+s.grossSales}/></>
          :selectedReport==="Cancel Items Summary"&&kotAudit?<DataGrid id="report-cancel-items" rows={kotAudit.cancellations} columns={[
            {key:"punchTime",label:"KOT punch time"},{key:"billNo",label:"Bill"},{key:"orderId",label:"Order ID"},{key:"kotId",label:"KOT ID"},{key:"item",label:"Item"},{key:"qty",label:"Qty",numeric:true},{key:"listedValue",label:"Original list value",numeric:true,render:x=>money(x.listedValue)},{key:"user",label:"User"},{key:"reason",label:"Reason",render:x=>x.reason||<Badge tone="bad">Missing</Badge>},{key:"status",label:"Status",render:x=><Badge tone="bad">{x.status}</Badge>}
          ]} totals={{billNo:"TOTAL",qty:num(kotAudit.summary.cancelledQuantity),listedValue:money(kotAudit.summary.cancelledListValue)}}/>
          :selectedReport==="Order Type Summary"?<DataGrid id="report-orders" rows={zOrderRows.map(x=>({...x,bills:x.orders,total:x.revenue,mix:x.revenue/+s.grossSales*100}))} columns={[{key:"name",label:"Order type"},{key:"bills",label:"Bills",numeric:true},{key:"guests",label:"Guests",numeric:true},{key:"gross",label:"Gross Item Sales",numeric:true,render:x=>money(x.gross)},{key:"discount",label:"Discounts",numeric:true,render:x=>money(x.discount)},{key:"complimentary",label:"Complimentary",numeric:true,render:x=>money(x.complimentary)},{key:"netItemsInclVat",label:"Net Item Sales incl. VAT",numeric:true,render:x=>money(x.netItemsInclVat)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"net",label:"Net Item Sales excl. VAT",numeric:true,render:x=>money(x.net)},{key:"charges",label:"Charges incl. VAT",numeric:true,render:x=>money(x.charges)},{key:"total",label:"Actual Sales incl. VAT",numeric:true,render:x=>money(x.total)},{key:"mix",label:"Mix %",numeric:true,render:x=>`${x.mix.toFixed(2)}%`}]} totals={{name:"TOTAL",bills:num(+s.fulfilledInvoices),guests:num(+s.totalGuests),gross:money(+s.correctedGrossBeforeDiscount),discount:money(+s.canonicalOrderDiscount),complimentary:money(+s.canonicalItemDiscount),netItemsInclVat:money(+s.taxableGrossInclVat),vat:money(+s.vat),net:money(+s.netSales),charges:money(+s.charges),total:money(+s.grossSales),mix:"100.00%"}}/>
          :<div className="noSource"><b>{selectedReport} layout is ready</b><p>The supplied exports do not contain the complete event-level source required to populate this report without inventing records. Its governed headers are:</p><div>{reportHeaders[selectedReport]?.map(h=><Badge key={h}>{h}</Badge>)}</div></div>}
          <Info title="Preview rule" tone="blue">Populated previews above use the current June dataset and the canonical discount calculation. Empty operational views are deliberately not fabricated; they need cash-drawer, expense, cancellation, wallet, receivable or settlement event facts.</Info>
        </section>
      </>}

      {tab==="menu"&&<>
        <section className="hero slim"><div><Badge tone="good">Editable planning model</Badge><h2>Menu engineering & price simulator</h2><p>Enter a food cost per item. Values are saved in this browser. Classification uses quantity popularity and contribution margin; it does not alter the audited sales ledger.</p></div></section>
        <section className="kpiGrid compact">
          {["Star","Plowhorse","Puzzle","Dog"].map(k=><Kpi key={k} label={`${k}s`} value={num(menuRows.filter(x=>x.classification===k).length)} sub={k==="Star"?"Popular · high margin":k==="Plowhorse"?"Popular · lower margin":k==="Puzzle"?"Lower popularity · high margin":"Lower popularity · lower margin"} tone={k==="Dog"?"bad":k==="Puzzle"?"warn":"good"}/>)}
        </section>
        <div className="menuControl"><label>Target food-cost percentage <input type="number" min="5" max="90" value={targetFoodCost} onChange={e=>setTargetFoodCost(+e.target.value||30)}/>%</label><span>Ideal selling price = food cost ÷ target food-cost %</span></div>
        <DataGrid id="menu" rows={menuRows} columns={[
          {key:"name",label:"Menu item"},{key:"qty",label:"Qty sold",numeric:true},{key:"net",label:"Realized sales",numeric:true,render:x=>money(x.net)},
          {key:"price",label:"Realized unit price",numeric:true,render:x=>money(x.price)},
          {key:"cost",label:"Food cost / unit",numeric:true,render:x=><input className="costInput" type="number" min="0" step=".01" value={costs[x.name]||""} placeholder="Enter cost" onClick={e=>e.stopPropagation()} onChange={e=>setCosts(c=>({...c,[x.name]:+e.target.value||0}))}/>},
          {key:"margin",label:"Margin / unit",numeric:true,render:x=>x.cost?money(x.margin):"Enter cost"},
          {key:"totalMargin",label:"Total contribution",numeric:true,render:x=>x.cost?money(x.totalMargin):"—"},
          {key:"idealPrice",label:"Ideal price",numeric:true,render:x=>x.cost?money(x.idealPrice):"—"},
          {key:"priceGap",label:"Price vs ideal",numeric:true,value:x=>x.cost?x.price-x.idealPrice:0,render:x=>x.cost?money(x.price-x.idealPrice):"—"},
          {key:"classification",label:"Class",render:x=><Badge tone={x.classification==="Dog"?"bad":x.classification==="Puzzle"?"warn":"good"}>{x.classification}</Badge>}
        ]} totals={{name:"TOTAL",qty:num(menuRows.reduce((a,x)=>a+x.qty,0)),net:money(menuRows.reduce((a,x)=>a+x.net,0)),totalMargin:money(menuRows.reduce((a,x)=>a+(x.cost?x.totalMargin:0),0))}}/>
        <Info title="Decision guide" tone="blue">Stars: protect availability and price. Plowhorses: improve portion cost or gently raise price. Puzzles: improve placement, naming or staff recommendations. Dogs: simplify, redesign or remove only after checking strategic value. Configure recipes and effective-dated ingredient costs before using this for accounting profit.</Info>
      </>}

      {tab==="guide"&&<>
        <section className="hero"><div><Badge tone="good">Embedded documentation</Badge><h2>How to read this audit dashboard</h2><p>Use controls, not labels, to decide whether a number is trustworthy. Every metric should state its grain, inclusion rule, tax basis and reconciliation target.</p></div></section>
        <section className="guideGrid">
          <article className="guideCard"><span>01</span><h3>Start at Gross Item Sales</h3><p>⃃ 118,066.00 preserves item value before Discounts and Complimentary.</p></article>
          <article className="guideCard"><span>02</span><h3>Reach Actual Sales</h3><p>⃃ 107,591.75 Net Item Sales incl. VAT + ⃃ 581.00 Charges incl. VAT = ⃃ 108,172.75 Actual Sales incl. VAT.</p></article>
          <article className="guideCard"><span>03</span><h3>Trace dimensions to lines</h3><p>Category, kitchen and item reports must allocate Discounts and classify Complimentary separately.</p></article>
          <article className="guideCard"><span>04</span><h3>Keep cancellations separate</h3><p>13 cancelled orders, 103 cancelled KOT items and post-sale refunds are different events.</p></article>
          <article className="guideCard"><span>05</span><h3>Explode split tenders</h3><p>1,831 payment allocations across 1,816 invoices is valid. Missing allocation amounts are not.</p></article>
          <article className="guideCard"><span>06</span><h3>Version menu prices</h3><p>Current-price differences become provable only when item IDs and effective-dated price versions are stored.</p></article>
        </section>
        <section className="panel"><div className="panelHead"><div><h3>Canonical equations</h3><p>These exact titles and equations apply to every dashboard, report and export.</p></div></div>
          <div className="formulaList">
            <div><code>Gross Item Sales</code><b>Σ(quantity × transaction-time unit price)</b></div>
            <div><code>Total Reductions</code><b>Discounts + Complimentary</b></div>
            <div><code>Net Item Sales incl. VAT</code><b>Gross Item Sales − Discounts − Complimentary</b></div>
            <div><code>Net Item Sales excl. VAT</code><b>Net Item Sales incl. VAT − item VAT</b></div>
            <div><code>Actual Sales incl. VAT</code><b>Net Item Sales incl. VAT + Charges incl. VAT + round-off</b></div>
            <div><code>Open due</code><b>Actual Sales incl. VAT − payments − wallet − credits</b></div>
          </div>
        </section>
        <Info title="Inventory warning" tone="red">Raw-material cost is zero because recipes and costs are not configured. ⃃ 118,053 is gross item subtotal—not gross margin. A 100% gross-profit percentage must not be displayed until COGS exists.</Info>
      </>}
    </main>
  </div>
}
