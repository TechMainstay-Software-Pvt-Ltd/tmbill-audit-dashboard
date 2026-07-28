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
const money = (v:number) => new Intl.NumberFormat("en-AE",{style:"currency",currency:"AED",minimumFractionDigits:2}).format(v);
const num = (v:number) => new Intl.NumberFormat("en-AE").format(v);

const reportSpecs = [
 ["Sales / Z Summary","Branch × business date × shift","The master financial control: gross sales, net ex VAT, VAT, discounts, charges, paid, due and variance."],
 ["Cash Audit","Drawer × cashier × shift","Opening float through counted cash, with every cash movement and variance."],
 ["Order Type Summary","Period × order type","Dine-in, pickup, delivery and aggregator mix using the same canonical invoice sales."],
 ["Payment Type Summary","Period × tender","One row per tender allocation, including every part of split payments."],
 ["Discount Summary","Rule × reason × user","Item and order discounts, eligibility, effective rate and approval exceptions."],
 ["Expense Summary","Expense transaction","Operational expense, input VAT, payment method, approval and reference."],
 ["Bill Summary","One row per invoice","The single invoice ledger from which all financial summaries are produced."],
 ["Delivery Boy Summary","Period × delivery employee","Orders, sales, delivery time, cash collections and handover variance."],
 ["Waiter Summary","Period × waiter","Bills, guests, average check, voids, tips and table turns."],
 ["Product Group Summary","Period × product group","Gross, discounts, net inclusive VAT, VAT, net ex VAT and mix."],
 ["Kitchen Department Summary","Period × kitchen","Production quantity, value, cancellations, preparation time and late tickets."],
 ["Category Summary","Period × category","Canonical line sales after allocated discounts; must sum to the item-sales control."],
 ["Sold Items Summary","Item × order type","Quantity, realized price, effective menu price, discounts, VAT and price variance."],
 ["Cancel Items Summary","One status event","Invoice/item/KOT identity, original value, actor, approver, reason and time."],
 ["Wallet Summary","One wallet event","Credits, debits, expiry, invoice reference and liability balance."],
 ["Due Payment Received","One receipt allocation","Receipt against original invoice with tender, reference and remaining balance."],
 ["Due Payment Receivable","One open invoice","Customer balance and aging from supplied but unpaid invoices."],
 ["Payment Variance","Period × tender/source","POS versus processor, bank and aggregator settlement reconciliation."],
 ["Currency Denominations","Drawer × denomination","Counted cash by denomination and verifier."],
 ["Order Source Summary","Period × source","Channel economics, cancellations, discounts, check average and settlement variance."]
];
const reportHeaders:Record<string,string[]> = {
 "Sales / Z Summary":["Business date","Shift","Bills","Gross items","Discount","Charges","VAT","Gross sales","Paid","Due","Variance"],
 "Cash Audit":["Shift","Drawer","Cashier","Opening float","Cash sales","Due received","Expenses","Expected","Counted","Variance"],
 "Order Type Summary":["Order type","Bills","Guests","Gross items","Discount","Charges","VAT","Gross sales","Average check","Mix %"],
 "Payment Type Summary":["Tender","Allocations","Invoices","Collected","Refunds","Net collected","Mix %","Variance"],
 "Discount Summary":["Discount","Level","Reason","User","Bills","Eligible gross","Discount","Effective %","Exceptions"],
 "Expense Summary":["Date","Expense ID","Category","Supplier","Net","Input VAT","Gross","Tender","Approver","Status"],
 "Bill Summary":["Bill","Order ID","Supply time","Status","Order type","Subtotal","Discount","Charges","VAT","Total","Paid","Due"],
 "Delivery Boy Summary":["Driver","Orders","Delivered","Cancelled","Sales","Cash collected","Due","Avg minutes","Late %","Variance"],
 "Waiter Summary":["Waiter","Bills","Guests","Sales","Discount","Average check","Voids","Tips","Table turns"],
 "Product Group Summary":["Product group","Qty","Gross","Item discount","Order discount","Net incl VAT","VAT","Net ex VAT","Mix %"],
 "Kitchen Department Summary":["Kitchen","Tickets","Lines","Qty","Net sales","Cancelled qty","Cancel value","Avg prep","Late %"],
 "Category Summary":["Category","Qty","Gross","Item discount","Order discount","Net incl VAT","VAT","Net ex VAT","Mix %"],
 "Sold Items Summary":["SKU","Item","Order type","Qty","Menu price","Actual price","Discount","Net","VAT","Price variance"],
 "Cancel Items Summary":["Event","Time","Invoice","KOT","Item","Qty","Original value","From","To","Reason","User","Approver"],
 "Wallet Summary":["Event","Time","Wallet","Type","Invoice","Credit","Debit","Expiry","Balance","Status"],
 "Due Payment Received":["Receipt","Received at","Invoice","Customer","Opening due","Received","Tender","Reference","Remaining"],
 "Due Payment Receivable":["Invoice","Supply date","Customer","Gross","Paid","Receipts","Credits","Open due","Age","Bucket"],
 "Payment Variance":["Date","Tender/source","POS net","Processor net","Fees","Expected deposit","Actual","Timing","Variance","Owner"],
 "Currency Denominations":["Date","Shift","Drawer","Currency","Denomination","Quantity","Amount","Counter","Verifier"],
 "Order Source Summary":["Source","Orders","Fulfilled","Cancelled","Discount","Charges","VAT","Sales","Average check","Cancel %","Mix %"]
};
const paymentMatrix=[
 {orderType:"Dine In",orders:684,card:29428.20,cash:5702.50,talabat:117,deliveroo:0,keeta:0,vat:1676.78,sales:35247.70,exported:34371.80},
 {orderType:"Pickup",orders:488,card:16713.45,cash:3549.40,talabat:0,deliveroo:0,keeta:0,vat:965.29,sales:20262.85,exported:19956.85},
 {orderType:"Delivery",orders:52,card:7179.50,cash:1478,talabat:0,deliveroo:0,keeta:0,vat:387,sales:8657.50,exported:8657.50},
 {orderType:"Talabat",orders:577,card:0,cash:0,talabat:42484.10,deliveroo:0,keeta:65.60,vat:2026.54,sales:42549.70,exported:42549.70},
 {orderType:"Keeta",orders:15,card:0,cash:0,talabat:0,deliveroo:0,keeta:1455,vat:69.31,sales:1455,exported:1455}
];
const sourceAudits=[
 {file:"daily-sales-report…xlsx",name:"Daily Sales Report",status:"fail",issue:"Export is effectively blank: only title, Discount, Total and an empty TOTAL row.",fix:"Populate governed Z-report sections from the invoice ledger; fail export generation when detail sections are empty.",headers:"DAY, DATE, BRANCH, DISCOUNT, TOTAL"},
 {file:"day-wise-consolidated-report…xlsx",name:"Day-wise Consolidated",status:"warn",issue:"Daily rows reconcile, but the workbook includes a grand-total row; naïve column sums double every value. Total Sale and Net Total are duplicate concepts.",fix:"Tag row_type, keep totals outside the data table, and define Net Sales as excluding VAT while Total Sale includes VAT and charges.",headers:"Order Date, Store Name, Total Sale, Sub Total, Discounts, Refunds, Refunded Tax, Net Sales, Taxes, Charges, Net Total, Tips, Total Orders, Average Per Order, Total Customers, Return Quantity, Round Off"},
 {file:"dsr-bill-no-of-series…xlsx",name:"Bill Number Series",status:"fail",issue:"All receipt-series fields are blank or zero despite fulfilled bill numbers #2313–#4141 and 1,816 bills.",fix:"Populate series start/end, issued count, void count, actual count and explicitly list sequence gaps.",headers:"State, Store Name, Starting No., Ending No., Count of Receipt No., Count of Void Invoices, Count of Actual Receipt"},
 {file:"dsr-bill-wise-report…xlsx",name:"DSR Bill-wise",status:"warn",issue:"Core sales and VAT reconcile, but Sale-5% is actually net sales excluding VAT; Net Sale repeats it. Delivery Charges show AED 520 while invoice Total Charges show AED 581.",fix:"Rename tax bases, add item/order discount separately, show charge type and tax, and retain one row_type-free invoice table.",headers:"Date, Time, Order Id, Ticket, Nature of Supply, Sale-5%, 5% VAT, Net Sale, Total Sale, Delivery Charges, Total Charges, Discount, Status, card, cash, deliveroo, keeta, talabat, Payment Mode, Guest"},
 {file:"dsr-day-wise…xlsx",name:"DSR Day-wise",status:"warn",issue:"Payment and sales totals reconcile, but Sale-5% and Net Sale repeat the same AED 103,047.83 and do not explain gross-before-discount or charge VAT.",fix:"Show original gross, item discount, order discount, taxable ex VAT, VAT, taxable incl VAT, charges and revenue as separate columns.",headers:"Date, Ticket count, card, cash, deliveroo, keeta, talabat, Sale-5%, VAT-5%, Net Sale, Delivery Charges, Round off, Total Sale"},
 {file:"dsr-item-wise…xlsx",name:"DSR Item-wise",status:"fail",issue:"Line Discount mixes allocated order discounts with true item discounts. #2990 exports Fattoush at price 0, discount 13 and negative net/tax base. Line net totals do not reconcile to invoice sales.",fix:"Add original unit price, item discount, allocated order discount, taxable incl/ex VAT and line VAT; prohibit negative taxable lines.",headers:"Date, Receipt IDs, Category, Item ID, Description, Sales Type, VAT %, Quantity, Price, Delivery Charges, Net Amount, Line Discount, Sale-Exempt, VAT Base Amount, Tax Product Group"},
 {file:"dsr-month-wise-report…xlsx",name:"DSR Month-wise",status:"warn",issue:"Core invoice totals reconcile, but Delivery/Total Charges are AED 520 instead of the invoice-ledger AED 581, leaving AED 61 unclassified.",fix:"Use the charge ledger, one column per charge type plus total charge VAT; rename Sale-5% to taxable sales excluding VAT.",headers:"Period, Tickets, Sale-5%, 5% VAT, Net Sale, Total Sale, Delivery Charges, Total Charges, Discount, payment columns"},
 {file:"month-wise-sales…xlsx",name:"Order Type Month-wise",status:"pass",issue:"Order-type subtotal, discount, VAT and total reconcile to the invoice ledger. It lacks charges, item discounts and validation columns.",fix:"Retain as a summary but add item/order discount split, charges, net ex VAT, VAT variance and reconciliation status.",headers:"Description, Sum of Sub Total, Sum of Total Discount, Sum of Actual VAT 5%, Sum of Total"},
 {file:"Morbido Express Restaurant Menu.xlsx",name:"Menu Master",status:"warn",issue:"Useful current menu snapshot, but no effective-from/to versioning and no immutable cross-report item key in every transaction export.",fix:"Version prices and tax groups by item ID, order type and effective dates; add recipe/food cost for menu engineering.",headers:"Short Code, Title, Tax Product Group, Category, Kitchen Dept, Base Item Price, status and platform fields"},
 {file:"Morbido Express Restaurant Multi Price File Format.xlsx",name:"Multi-price Master",status:"warn",issue:"Order-type prices exist, but only as a current snapshot. Historical June mismatches cannot be proven wrong from this file.",fix:"Store immutable item ID, order type, price, tax inclusion and effective-from/to dates.",headers:"Short Code, Item Name, Base Item Price, Quick Bill, Dine In, Pickup, Delivery, Talabat, Deliveroo, Keeta"},
 {file:"order-type-day-wise-report…xlsx",name:"Order Type Day-wise",status:"fail",issue:"Delivery rows violate Subtotal − Discount = Total because charges are embedded inconsistently; period grand subtotal/discount/total do not match the invoice ledger.",fix:"Keep item sales and charges separate, include VAT, and reconcile every order-type/day row to invoice IDs.",headers:"Date, Day, per-order-type Sub Total, Total Discount, Total, Grand Totals"},
 {file:"sales-report…xlsx",name:"Sales Report",status:"warn",issue:"Best invoice ledger and core control source, but it contains 12 identically named Delivery Charges headers, hiding charge identity. Tax On Charges is zero while Total Charges are AED 581.",fix:"Unique charge_type columns or a child charge table; show item/order discounts separately and validate charge tax treatment.",headers:"Order ID, Bill, dates, items, Sub Total, discount rules, Total Discount, VAT, tax modes, Total Tax, repeated charges, Item Level Charges, Total Charges, Tax On Charges, Total, Status, Type, Payment, Source, Guests"},
 {file:"sales-report-with-items…xlsx",name:"Sales Report With Items",status:"fail",issue:"Hierarchical invoice and item rows share columns and include a summary row, making flat sums duplicate invoice totals and obscuring discount allocation.",fix:"Export separate Invoice and Invoice Line sheets joined by Order ID; lines must carry both discount levels and VAT.",headers:"Invoice financial headers followed by HSN, Item Name, Price, Qty, line Total, Product Group, Category, Return Item"},
 {file:"sales-report-with-product-group-details…xlsx",name:"Product Group Details",status:"fail",issue:"Multiple columns are all named Delivery Charges; summary/footer rows sit inside the data range; Taxable Amount and VAT summary labels are inconsistent.",fix:"One normalized product-group table with unique headers, row_type, corrected allocated discounts, line VAT and explicit charge allocation policy.",headers:"Bill Date, Order ID, Bill, Qty, Subtotal, Taxable Amount, Discount, VAT 5%, repeated Delivery Charges, Round Off, Total Amount, 0%, Status"},
 {file:"simplified-day-wise-dsr…xlsx",name:"Simplified Day-wise DSR",status:"warn",issue:"Daily and period totals broadly reconcile, but Total Taxable Amount is gross pre-discount, not the taxable amount used for VAT. Credit combines aggregator tenders.",fix:"Rename gross pre-discount, show taxable base after discounts, split aggregator tenders and show charges/tax separately.",headers:"Day, Date, Invoice count, Total Taxable Amount, Non Taxable, Total Tax, tender sales/tax/ex-tax, Discount, Charges, Grand Total"},
 {file:"tax-submission-payment-report…xlsx",name:"Tax Submission Payment-wise",status:"fail",issue:"Total Collection is AED 106,990.85, short of invoice/payment sales by AED 1,181.90. Duplicate VAT headers make the grain ambiguous.",fix:"Build from payment-allocation facts, not a primary payment label; one row per invoice × tender with allocated VAT only for analytics, never as the filing basis.",headers:"Order Type, Order Count, repeated VAT columns, Card/Cash/Talabat/Deliveroo/Keeta/Due collections, Total Collection"},
 {file:"tax-submission-report…xlsx",name:"Tax Submission Bill-wise",status:"fail",issue:"Tender sales total only AED 106,990.85 while Total Sales is AED 108,172.75. Payment columns omit secondary split allocations although invoice VAT total is correct.",fix:"Separate supply/VAT filing ledger from payment reconciliation. Explode every tender allocation and validate tender sum = invoice total.",headers:"Bill Number, Order Type, per-tender VAT and Sales, Total VAT, Total Sales"},
 {file:"day-wise-summary-report…xlsx",name:"New Day-wise Summary",status:"pass",issue:"This is the strongest day-level control received: it reconciles 1,816 bills, AED 118,053 item subtotal, AED 10,461.25 header discounts, AED 5,124.92 VAT, AED 581 charges and AED 108,172.75 revenue. It omits at least the confirmed AED 13 item adjustment on #2990; completeness cannot be certified from this summary.",fix:"Retain its daily bridge but add original transaction gross, separately stored item discounts, order discounts and a completeness control sourced from immutable invoice lines.",headers:"Outlet, Date, Bill Range, Total Bill, Total Discount, Total Tax, Total Charges, Item Total, Net Sales, Grand Total"},
 {file:"discount-report…xlsx",name:"New Discount Report",status:"warn",issue:"Order-type discount total AED 10,461.25 reconciles to invoice header discounts, but the report has no discount rule, bill, item/order level, approver or eligibility fields. It omits the confirmed AED 13 item adjustment and cannot reveal partial item discounts.",fix:"Add invoice and item grain, original transaction-time price, discount level, eligible base, rule, reason, user and approval. Treat AED 10,474.25 as a confirmed minimum until line-level completeness is available.",headers:"Order From, Discount Amount, Status"},
 {file:"item-wise-report…xlsx",name:"New Item-wise Report",status:"fail",issue:"All 3,367 item rows report Discount = zero. Bill #2990 shows Fattoush price zero and discount zero, so the confirmed AED 13 item adjustment and all allocated order discounts disappear. Other partial item discounts cannot be distinguished from price-version changes.",fix:"Export original transaction-time price, item discount, allocated order discount, sold incl/ex VAT, line VAT and price-version ID. Join by immutable Order ID and item-line ID.",headers:"Order ID, Title, Quantity, Price, Discount Name, Discount, Discount Reason, Item Note, Bill Number, Order Date, Table Name"},
 {file:"order-state-transition-report…xlsx",name:"Order State Transition",status:"warn",issue:"Only 1,229 offline-type records are present (Dine In/Pickup/Delivery); all aggregator orders are absent. Food Ready and Dispatched timestamps are entirely empty, and some rows contain both completed and canceled timestamps.",fix:"Include every order source, use one immutable status-event row per transition, and validate mutually exclusive terminal states.",headers:"Order ID, Platform, Order Type, Store, Placed, Acknowledged, Food Ready, Dispatched, Completed, Canceled, Username, Source, Duration"},
 {file:"shift-wise-report…xlsx",name:"Shift-wise Report",status:"fail",issue:"Report header dates incorrectly show 13 Feb 2004 to 03 Apr 2006. Seven June shifts have zero Total Sale and the shift total is only AED 85,346.45 rather than AED 108,172.75.",fix:"Repair date formatting/source parameters; allocate every invoice to a shift by settlement timestamp and require shift totals to reconcile to day sales.",headers:"Shift Start, Start User, Shift End, End User, Closing Balance, Opening Balance, Expense, Total Sale, Current Closing Balance, Comments"},
 {file:"start-close-day-report…xlsx",name:"Start / Close Day Report",status:"fail",issue:"June sales total reconciles, but invoice ranges do not align with the actual June bill series and 16 days have non-zero cash differences totaling AED -10,189.20, often because Close Day Amount is zero.",fix:"Separate expected cash from counted cash, require count/approval before close, reconcile invoice range to issued bills and show explained versus unexplained variance.",headers:"Start Day, End Day, users, Opening/Closing Balance, Expense, Total Sale, Invoice Range, Comments, Close Day Amount, Cash Difference"},
];

function Badge({tone="neutral",children}:{tone?:string;children:React.ReactNode}) {
  return <span className={`badge ${tone}`}>{children}</span>
}
function Kpi({label,value,sub,tone="good"}:{label:string;value:string;sub:string;tone?:string}) {
  return <article className={`kpi ${tone}`}><div className="kpiLabel">{label}</div><div className="kpiValue">{value}</div><div className="kpiSub">{sub}</div></article>
}
function SalesReconciliation({itemSales,charges,actualSales}:{itemSales:number;charges:number;actualSales:number}) {
  const reconciled=+(itemSales+charges).toFixed(2);
  const variance=+(reconciled-actualSales).toFixed(2);
  return <div className="reportReconciliation" aria-label="Report to actual sales reconciliation">
    <div><span>Item sales after discounts · incl. VAT</span><b>{money(itemSales)}</b></div><i>+</i>
    <div><span>Invoice charges</span><b>{money(charges)}</b></div><i>=</i>
    <div className="reconciledTotal"><span>Reconciled actual sales</span><b>{money(reconciled)}</b></div>
    <div className={Math.abs(variance)<=.01?"reconPass":"reconFail"}><span>Variance to invoice sales</span><b>{money(variance)}</b><small>{Math.abs(variance)<=.01?"Matched":"Review required"}</small></div>
  </div>
}
function Info({title,children,tone="blue"}:{title:string;children:React.ReactNode;tone?:string}) {
  return <div className={`info ${tone}`}><strong>{title}</strong><div>{children}</div></div>
}
function Bar({value,max,color="#17a673"}:{value:number;max:number;color?:string}) {
  return <div className="bar"><i style={{width:`${Math.max(1,value/max*100)}%`,background:color}} /></div>
}
function TenderValue({row,tender}:{row:any;tender:"card"|"cash"|"talabat"|"deliveroo"|"keeta"}) {
  const value=Number(row[tender]||0);
  if(!value)return <span className="zeroDash">—</span>;
  const aggregator=["Talabat","Deliveroo","Keeta"].includes(row.orderType);
  const invalid=aggregator?tender.toLowerCase()!==row.orderType.toLowerCase():["talabat","deliveroo","keeta"].includes(tender);
  return <span className={invalid?"invalidTender":""}>{money(value)}{invalid&&<small>Unexpected channel</small>}</span>
}
type GridCol={key:string;label:string;numeric?:boolean;value?:(r:any)=>string|number;render?:(r:any)=>React.ReactNode};
function DataGrid({id,rows,columns,totals,onRowClick,selectedKey}:{id:string;rows:any[];columns:GridCol[];totals?:Record<string,React.ReactNode>;onRowClick?:(r:any)=>void;selectedKey?:string}) {
  const [order,setOrder]=useState(columns.map(c=>c.key));
  const [visible,setVisible]=useState<Record<string,boolean>>(()=>Object.fromEntries(columns.map(c=>[c.key,true])));
  const [sort,setSort]=useState<{key:string;dir:1|-1}|null>(null);
  const [chooser,setChooser]=useState(false);
  const colMap=useMemo(()=>new Map(columns.map(c=>[c.key,c])),[columns]);
  const active=order.filter(k=>visible[k]).map(k=>colMap.get(k)!).filter(Boolean);
  const sorted=useMemo(()=>sort?[...rows].sort((a,b)=>{const c=colMap.get(sort.key);const av=c?.value?c.value(a):a[sort.key];const bv=c?.value?c.value(b):b[sort.key];return (typeof av==="number"&&typeof bv==="number"?av-bv:String(av??"").localeCompare(String(bv??"")))*sort.dir}):rows,[rows,sort,colMap]);
  const move=(from:string,to:string)=>setOrder(prev=>{const n=prev.filter(x=>x!==from);n.splice(n.indexOf(to),0,from);return n});
  return <div className={`gridShell dataGrid-${id}`}><div className="gridTools"><span>{num(rows.length)} rows</span><button type="button" onClick={e=>{e.stopPropagation();setChooser(v=>!v)}}>Columns ▾</button>{chooser&&<div className="columnChooser" role="group" aria-label="Visible table columns" onClick={e=>e.stopPropagation()}><strong>Show or hide columns</strong>{columns.map(c=><label key={c.key} onClick={e=>e.stopPropagation()}><input type="checkbox" checked={visible[c.key]!==false} onClick={e=>e.stopPropagation()} onChange={e=>setVisible(v=>({...v,[c.key]:e.target.checked}))}/><span>{c.label}</span></label>)}</div>}</div>
    <div className="tablePanel smartGrid"><table><thead><tr>{active.map(c=><th key={c.key} draggable onDragStart={e=>e.dataTransfer.setData("text/plain",c.key)} onDragOver={e=>e.preventDefault()} onDrop={e=>move(e.dataTransfer.getData("text/plain"),c.key)} onClick={()=>setSort(s=>s?.key===c.key?{key:c.key,dir:s.dir===1?-1:1}:{key:c.key,dir:1})} className={c.numeric?"numeric":""}>{c.label}<i>{sort?.key===c.key?(sort.dir===1?" ↑":" ↓"):" ↔"}</i></th>)}</tr></thead>
      <tbody>{sorted.map((r,i)=><tr key={r.id||r.rowId||r.name||i} onClick={()=>onRowClick?.(r)} className={(selectedKey&&(r.billNo===selectedKey||r.id===selectedKey))?"selected":""}>{active.map(c=><td key={c.key} className={c.numeric?"numeric":""}>{c.render?c.render(r):String(c.value?c.value(r):r[c.key]??"—")}</td>)}</tr>)}</tbody>
      {totals&&<tfoot><tr>{active.map(c=><td key={c.key} className={c.numeric?"numeric":""}>{totals[c.key]??""}</td>)}</tr></tfoot>}</table></div></div>
}

export default function Home() {
  const [data,setData]=useState<AuditData|null>(null);
  const [paymentRecon,setPaymentRecon]=useState<PaymentReconciliation|null>(null);
  const [separateAudit,setSeparateAudit]=useState<SeparateAudit|null>(null);
  const [kotAudit,setKotAudit]=useState<KotAudit|null>(null);
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
  useEffect(()=>{fetch("/data/audit-data.json").then(r=>r.json()).then(setData)},[]);
  useEffect(()=>{fetch("/data/payment-reconciliation.json").then(r=>r.json()).then(setPaymentRecon)},[]);
  useEffect(()=>{fetch("/data/separate-sales-audit.json").then(r=>r.json()).then(setSeparateAudit)},[]);
  useEffect(()=>{fetch("/data/item-kot-audit.json").then(r=>r.json()).then(setKotAudit)},[]);
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
  return <div className="app">
    <aside className="sidebar">
      <div className="brand"><div className="logo">TM</div><div><b>TMBill Audit</b><small>Revenue intelligence</small></div></div>
      <div className="period"><span>Review period</span><b>June 2026</b><small>The SS · Morbido Express</small></div>
      <nav>{tabs.map(([id,label])=><button key={id} className={tab===id?"active":""} onClick={()=>setTab(id)}><span>{label}</span>{id==="bills"&&<em>1</em>}{id==="items"&&<em>915</em>}</button>)}</nav>
      <div className="sideControl"><span>Overall reconciliation</span><strong>3 confirmed defects</strong><small>Core invoice ledger passes</small></div>
    </aside>
    <main className="main">
      <header className="topbar">
        <div><div className="eyebrow">THE SS · MORBIDO EXPRESS RESTAURANT</div><h1>{tabs.find(x=>x[0]===tab)?.[1]}</h1></div>
        <div className="topActions"><span className="sync"><i/> Data frozen at 30 Jun 2026, 11:58 PM</span><button className="helpBtn" onClick={()=>setTab("guide")}>Metric guide</button></div>
      </header>

      {tab==="overview"&&<>
        <section className="hero">
          <div><Badge tone="good">Invoice ledger reconciled</Badge><h2>Sales are right. Some dimensions are not.</h2><p>The fulfilled invoice ledger ties across reports. The red controls below identify where item, payment and category calculations break after invoice settlement.</p></div>
          <div className="heroScore"><b>97%</b><span>financial control confidence</span><small>3 exceptions require correction</small></div>
        </section>
        <section className="kpiGrid">
          <Kpi label="Original item gross" value={money(+s.correctedGrossBeforeDiscount)} sub={`${num(+s.totalItemQuantity)} items on ${num(+s.fulfilledInvoices)} fulfilled bills`}/>
          <Kpi label="Confirmed item-level discount" value={money(+s.canonicalItemDiscount)} sub="Minimum confirmed: #2990 Fattoush; partial line discounts are not provable from exports"/>
          <Kpi label="Order-level discounts" value={money(+s.canonicalOrderDiscount)} sub="Allocated only across eligible lines"/>
          <Kpi label="Total discounts" value={money(+s.canonicalTotalDiscount)} sub="Item + order discounts; fully reconciled"/>
          <Kpi label="Taxable total incl VAT" value={money(+s.taxableGrossInclVat)} sub="Original gross less all discounts"/>
          <Kpi label="Output VAT" value={money(+s.vat)} sub="Canonical reconstruction AED 5,124.93"/>
          <Kpi label="Total revenue" value={money(+s.grossSales)} sub="Taxable items + AED 581.00 charges"/>
          <Kpi label="Bill series" value={`#${s.billFrom}–#${s.billTo}`} sub={`${num(+s.fulfilledInvoices)} fulfilled bills generated`}/>
          <Kpi label="Average cheque" value={money(+s.averageCheck)} sub="Revenue per fulfilled bill"/>
          <Kpi label="Average per person" value={money(+s.averagePerPerson)} sub={`${num(+s.totalGuests)} recorded guests`}/>
          <Kpi tone="bad" label="Payment allocation gap" value={money(+s.paymentComponentGap)} sub="Missing secondary split-tender allocations"/>
          <Kpi tone="warn" label="Charge classification gap" value={money(61)} sub="All charges 581 vs DSR delivery charges 520"/>
          <Kpi tone="bad" label="Category overstatement" value={money(+s.categoryOverstatement)} sub="Item export net exceeds invoice total"/>
        </section>
        <section className="twoCol">
          <div className="panel"><div className="panelHead"><div><h3>Sales by order type</h3><p>These values reconcile exactly to gross sales.</p></div><Badge tone="good">PASS</Badge></div>
            <div className="rankList">{orderTotals.map(([name,value,count])=><div className="rank" key={name}><div><b>{name}</b><span>{count} orders</span></div><div className="grow"><Bar value={value} max={42549.7}/></div><strong>{money(value)}</strong></div>)}</div>
          </div>
          <div className="panel"><div className="panelHead"><div><h3>Issue trace</h3><p>Start with the largest control failures.</p></div></div>
            <button className="issueRow critical" onClick={()=>{setSelectedBill("2990");setTab("bills")}}><span>01</span><div><b>Malformed item line</b><small>Bill 2990 · Fattoush · negative taxable base</small></div><strong>Trace →</strong></button>
            <button className="issueRow high" onClick={()=>setTab("payments")}><span>02</span><div><b>Incomplete split payments</b><small>Tax payment components short by AED 1,181.90</small></div><strong>Trace →</strong></button>
            <button className="issueRow medium" onClick={()=>setTab("categories")}><span>03</span><div><b>Category denominator mismatch</b><small>Percentages add to 106.29%, not 100%</small></div><strong>Trace →</strong></button>
          </div>
        </section>
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
            {key:"billNo",label:"Bill",render:x=><><b>#{x.billNo}</b><small>{x.id}</small></>},{key:"date",label:"Date"},{key:"orderType",label:"Order type"},{key:"guests",label:"Guests",numeric:true},
            {key:"canonicalGrossBeforeDiscount",label:"Original gross",numeric:true,render:x=>money(x.canonicalGrossBeforeDiscount)},
            {key:"canonicalItemDiscount",label:"Item discount",numeric:true,render:x=>money(x.canonicalItemDiscount)},
            {key:"canonicalOrderDiscount",label:"Order discount",numeric:true,render:x=>money(x.canonicalOrderDiscount)},
            {key:"canonicalTotalDiscount",label:"Total discount",numeric:true,render:x=>money(x.canonicalTotalDiscount)},
            {key:"canonicalTaxableInclVat",label:"Net incl VAT",numeric:true,render:x=>money(x.canonicalTaxableInclVat)},
            {key:"canonicalNetExVat",label:"Net ex VAT",numeric:true,render:x=>money(x.canonicalNetExVat)},
            {key:"canonicalVat",label:"VAT",numeric:true,render:x=>money(x.canonicalVat)},{key:"charges",label:"Charges",numeric:true,render:x=>money(x.charges)},{key:"total",label:"Revenue",numeric:true,render:x=><b>{money(x.total)}</b>},
            {key:"discountAssessment",label:"Discount check",render:x=><Badge tone={x.discountAssessment==="corrected"?"warn":"good"}>{x.discountAssessment}</Badge>},
            {key:"vatAssessment",label:"VAT check",render:x=><Badge tone={x.vatAssessment==="exact"?"good":"warn"}>{x.vatAssessment}</Badge>},
            {key:"controlStatus",label:"Control",render:x=><Badge tone={x.controlStatus==="pass"?"good":x.controlStatus==="corrected"?"warn":"bad"}>{x.controlStatus}</Badge>}
          ]} totals={{billNo:"TOTAL",guests:num(filteredInvoices.reduce((a,x)=>a+x.guests,0)),canonicalGrossBeforeDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalGrossBeforeDiscount,0)),canonicalItemDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalItemDiscount,0)),canonicalOrderDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalOrderDiscount,0)),canonicalTotalDiscount:money(filteredInvoices.reduce((a,x)=>a+x.canonicalTotalDiscount,0)),canonicalTaxableInclVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalTaxableInclVat,0)),canonicalNetExVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalNetExVat,0)),canonicalVat:money(filteredInvoices.reduce((a,x)=>a+x.canonicalVat,0)),charges:money(filteredInvoices.reduce((a,x)=>a+x.charges,0)),total:money(filteredInvoices.reduce((a,x)=>a+x.total,0))}}/>
          <div className="drawer">{selected&&<><div className="drawerHead"><div><span>INVOICE TRACE</span><h3>Bill #{selected.billNo}</h3><p>{selected.date} · {selected.time} · {selected.orderType}</p></div>{selected.hasAnomaly?<Badge tone="bad">FAIL</Badge>:<Badge tone="good">PASS</Badge>}</div>
            <div className="bridge"><div><span>Subtotal</span><b>{money(selected.subtotal)}</b></div><i>−</i><div><span>Invoice discount</span><b>{money(selected.discount)}</b></div><i>+</i><div><span>Charges</span><b>{money(selected.charges)}</b></div><i>=</i><div className="total"><span>Gross total</span><b>{money(selected.total)}</b></div></div>
            <div className="controlBox bad"><span>ITEM ↔ INVOICE DISCOUNT CONTROL</span><div><b>Item lines</b><strong>{money(selected.itemLineDiscount)}</strong></div><div><b>Invoice header</b><strong>{money(selected.discount)}</strong></div><div><b>Variance</b><strong>{money(selected.discountVariance)}</strong></div></div>
            <div className="controlBox"><span>CORRECTED DISCOUNT & VAT BRIDGE</span><div><b>Original item gross</b><strong>{money(selected.canonicalGrossBeforeDiscount)}</strong></div><div><b>Item-level discount</b><strong>− {money(selected.canonicalItemDiscount)}</strong></div><div><b>Eligible order discount</b><strong>− {money(selected.canonicalOrderDiscount)}</strong></div><div><b>Taxable incl VAT</b><strong>{money(selected.canonicalTaxableInclVat)}</strong></div><div><b>VAT (5/105)</b><strong>{money(selected.canonicalVat)}</strong></div><div><b>Net ex VAT</b><strong>{money(selected.canonicalNetExVat)}</strong></div></div>
            <h4>Item lines</h4>{selectedLines.map(x=><div className={`lineCard ${x.anomaly?"anomaly":""}`} key={x.rowId}><div><b>{x.name}</b><small>{x.itemId} · Qty {x.qty} · {x.category}</small></div><div><span>Price</span><b>{money(x.actualPrice)}</b></div><div><span>Discount</span><b>{money(x.lineDiscount)}</b></div><div><span>Net export</span><b>{money(x.netAmount)}</b></div>{x.anomaly&&<Badge tone="bad">Negative taxable line</Badge>}</div>)}
            {selected.billNo==="2990"&&<Info title="Correct treatment for #2990" tone="red">Fattoush original value AED 13.00 is fully removed as an item discount. It is excluded from the 10% order-discount base. The eligible AED 81.00 receives AED 8.10 order discount: original gross AED 94.00 − item discount AED 13.00 − order discount AED 8.10 = taxable total AED 72.90, VAT AED 3.47 and net ex VAT AED 69.43.</Info>}
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
          {key:"monthlyAvgGrossUnit",label:"Monthly avg list",numeric:true,render:x=>money(x.monthlyAvgGrossUnit)},{key:"orderTypeAvgGrossUnit",label:"Order-type avg list",numeric:true,render:x=>money(x.orderTypeAvgGrossUnit)},
          {key:"canonicalItemDiscount",label:"Item discount",numeric:true,render:x=>money(x.canonicalItemDiscount)},{key:"canonicalOrderDiscount",label:"Order discount",numeric:true,render:x=>money(x.canonicalOrderDiscount)},{key:"effectiveDiscountRate",label:"Effective discount %",numeric:true,render:x=>`${x.effectiveDiscountRate.toFixed(2)}%`},
          {key:"canonicalNetInclVat",label:"Sold incl VAT",numeric:true,render:x=>money(x.canonicalNetInclVat)},{key:"canonicalNetExVat",label:"Sold ex VAT",numeric:true,render:x=>money(x.canonicalNetExVat)},{key:"canonicalVat",label:"VAT",numeric:true,render:x=>money(x.canonicalVat)},
          {key:"monthlyAvgSoldInclVat",label:"Monthly avg sold incl VAT",numeric:true,render:x=>money(x.monthlyAvgSoldInclVat)},{key:"monthlyAvgSoldExVat",label:"Monthly avg sold ex VAT",numeric:true,render:x=>money(x.monthlyAvgSoldExVat)},
          {key:"orderTypeAvgSoldInclVat",label:"Order-type avg sold incl VAT",numeric:true,render:x=>money(x.orderTypeAvgSoldInclVat)},{key:"orderTypeAvgSoldExVat",label:"Order-type avg sold ex VAT",numeric:true,render:x=>money(x.orderTypeAvgSoldExVat)},
          {key:"priceStatus",label:"Menu comparison",render:x=><Badge tone={x.anomaly?"bad":x.priceStatus==="match"?"good":x.priceStatus==="mismatch"?"warn":"neutral"}>{x.anomaly?"Data anomaly":x.priceStatus}</Badge>}
        ]} totals={{name:"TOTAL",qty:num(filteredItems.reduce((a,x)=>a+x.qty,0)),canonicalItemDiscount:money(filteredItems.reduce((a,x)=>a+x.canonicalItemDiscount,0)),canonicalOrderDiscount:money(filteredItems.reduce((a,x)=>a+x.canonicalOrderDiscount,0)),canonicalNetInclVat:money(filteredItems.reduce((a,x)=>a+x.canonicalNetInclVat,0)),canonicalNetExVat:money(filteredItems.reduce((a,x)=>a+x.canonicalNetExVat,0)),canonicalVat:money(filteredItems.reduce((a,x)=>a+x.canonicalVat,0))}}/>}
        {tab==="items"&&<Info title="Price comparison limitation" tone="amber">The menu file is a current snapshot. Without effective-from/effective-to dates, a mismatch is an audit exception—not proof that the June transaction was incorrectly priced. Production should join by immutable item ID and the price version effective at supply time.</Info>}
      </section>}

      {tab==="discounts"&&<>
        <section className="kpiGrid compact"><Kpi label="Invoice discounts" value={money(10461.25)} sub="Canonical header total"/><Kpi label="Item-export discounts" value={money(10474.25)} sub="Includes malformed bill 2990 line" tone="bad"/><Kpi label="Variance" value={money(13)} sub="Confirmed defect" tone="bad"/><Kpi label="Discounted invoices" value="587" sub="250 offline · 337 online"/></section>
        <section className="panel"><div className="panelHead"><div><h3>Required discount allocation waterfall</h3><p>The exact calculation every category and item report should use.</p></div></div>
          <div className="steps">{[
            ["1","Line gross","Quantity × transaction unit price"],
            ["2","Item discount","Subtract explicit line-level discount"],
            ["3","Eligible base","Exclude non-discountable lines and partition by tax rate"],
            ["4","Allocate order discount","Order discount × eligible line gross ÷ eligible invoice gross"],
            ["5","Absorb rounding residual","Last eligible line receives the AED 0.01 residual"],
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
        <Info title="New confirmation for bill #2990" tone="amber">The Complimentary Items report independently records Fattoush, quantity 1, AED 13.00, created by tarekbmr on 12 Jun 2026 at 08:41:17 with reason “nnnno tomatto”. The KOT report has the same item/date/user/reason on #2990, settled at zero, so this is an exact contextual match. However, the complimentary report uses a different numeric KOT reference; developers must export the immutable order ID, bill number, KOT ID and line ID directly.</Info>
        <div className="filters kotFilters"><label className="search"><span>⌕</span><input value={kotQuery} onChange={e=>setKotQuery(e.target.value)} placeholder="Search bill, order ID, KOT, item, table, user or reason…"/></label>
          <select value={kotStatus} onChange={e=>setKotStatus(e.target.value)}><option>All</option><option>Placed</option><option>Cancelled</option><option>Edited</option></select>
          <select value={kotUser} onChange={e=>setKotUser(e.target.value)}><option>All</option>{Object.keys(kotAudit.summary.users).map(x=><option key={x}>{x}</option>)}</select>
          <select value={kotDate} onChange={e=>setKotDate(e.target.value)}><option>All</option>{[...new Set(kotAudit.events.map(x=>x.date))].sort().map(x=><option key={x}>{x}</option>)}</select>
          <span className="resultCount">{num(filteredKotEvents.length)} events</span>
        </div>
        <section className="kotWorkspace">
          <div>
            <DataGrid id="kot-events" rows={pagedKotEvents} selectedKey={selectedKotBill} onRowClick={x=>setSelectedKotBill(x.billNo)} columns={[
              {key:"billNo",label:"Bill",render:x=><><b>#{x.billNo}</b><small>{x.orderId}</small></>},{key:"date",label:"Bill date"},{key:"billTime",label:"Closed"},{key:"punchTime",label:"KOT punched"},
              {key:"kotNo",label:"KOT no."},{key:"item",label:"Item"},{key:"qty",label:"Punched qty",numeric:true},{key:"settledQty",label:"Settled qty",numeric:true},
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
        <section className="twoCol"><div className="panel"><div className="panelHead"><div><h3>Correct payment summary</h3><p>Dashboard and printed report.</p></div><Badge tone="good">PASS</Badge></div>{paymentTotals.map(([name,value,count])=><div className="rank" key={name}><div><b>{name}</b><span>{count} allocations</span></div><div className="grow"><Bar value={value} max={53321.15}/></div><strong>{money(value)}</strong></div>)}</div>
          <div className="panel"><div className="panelHead"><div><h3>Broken tax-payment export</h3><p>Payment component bridge.</p></div><Badge tone="bad">FAIL</Badge></div>
            {[["Card",52998.35],["Cash",9870.8],["Talabat",42601.1],["Keeta",1520.6]].map(([n,v])=><div className="ledgerRow" key={n as string}><span>{n}</span><b>{money(v as number)}</b></div>)}
            <div className="ledgerTotal"><span>Component total</span><b>{money(106990.85)}</b></div><div className="ledgerGap"><span>Missing allocation</span><b>{money(1181.9)}</b></div>
          </div></section>
        <Info title="Correct data model" tone="blue">Create one payment-allocation row per invoice × tender. A split invoice must have multiple rows. Payment totals must come from allocation amounts; the invoice’s primary payment label is not sufficient.</Info>
      </>}

      {tab==="crosscheck"&&<>
        <section className="hero slim"><div><Badge tone="bad">TMBill versus canonical controls</Badge><h2>See the reported number, correct number and root cause together.</h2><p>Sales/order type is a supply dimension. Payment type is a settlement dimension. The AED 1,181.90 was collected correctly on 15 split bills but omitted from the Tax Submission report’s secondary-tender columns.</p></div><div className="heroScore bad"><b>{money(1181.90)}</b><span>split tender omitted from tax report</span><small>Customer payments are fully reconciled</small></div></section>
        <section className="kpiGrid compact downloadableKpis">
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="Correct invoice sales" value={money(108172.75)} sub="Supply ledger · click to download Excel"/></a>
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="TMBill payment-tax export" value={money(106990.85)} sub="Incomplete allocations · download detail" tone="bad"/></a>
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="Confirmed discount correction" value={money(13)} sub="Bill #2990 · download control" tone="warn"/></a>
          <a href="/downloads/TMBill_Reconciliation_Control.xlsx" download><Kpi label="Unclassified charge gap" value={money(61)} sub="AED 581 vs AED 520 · download detail" tone="bad"/></a>
        </section>
        <section className="panel"><div className="panelHead"><div><h3>Order type × payment type × VAT</h3><p>Correct matrix rebuilt from the bill-wise payment allocations. Click headers to sort or Columns to configure.</p></div></div>
          <DataGrid id="order-payment-matrix" rows={paymentMatrix.map(x=>({...x,variance:x.sales-x.exported,validation:(x.orderType==="Talabat"&&x.keeta>0)||(x.orderType==="Dine In"&&x.talabat>0)?"Channel mismatch":Math.abs(x.sales-x.exported)>.01?"Missing allocation":"Pass"}))} columns={[
            {key:"orderType",label:"Order type",render:x=><b>{x.orderType}</b>},{key:"orders",label:"Orders",numeric:true},{key:"card",label:"Card",numeric:true,render:x=><TenderValue row={x} tender="card"/>},{key:"cash",label:"Cash",numeric:true,render:x=><TenderValue row={x} tender="cash"/>},{key:"talabat",label:"Talabat",numeric:true,render:x=><TenderValue row={x} tender="talabat"/>},{key:"deliveroo",label:"Deliveroo",numeric:true,render:x=><TenderValue row={x} tender="deliveroo"/>},{key:"keeta",label:"Keeta",numeric:true,render:x=><TenderValue row={x} tender="keeta"/>},{key:"sales",label:"Correct sales",numeric:true,render:x=>money(x.sales)},{key:"vat",label:"Invoice VAT",numeric:true,render:x=>money(x.vat)},{key:"exported",label:"Tax report captured",numeric:true,render:x=>money(x.exported)},{key:"variance",label:"Split tender omitted",numeric:true,render:x=>x.variance?money(x.variance):"—"},{key:"validation",label:"Validation",render:x=><Badge tone={x.validation==="Pass"?"good":"bad"}>{x.validation}</Badge>}
          ]} totals={{orderType:"TOTAL",orders:"1,816",card:money(53321.15),cash:money(10729.90),talabat:money(42601.10),deliveroo:money(0),keeta:money(1520.60),sales:money(108172.75),vat:money(5124.92),exported:money(106990.85),variance:money(1181.90)}}/>
        </section>
        <Info title="What the reconciliation means" tone="red"><ul><li><b>AED 875.90</b> is the secondary Card/Cash portion of 11 Dine In split bills omitted by the Tax Submission report.</li><li><b>AED 306.00</b> is the secondary Card/Cash portion of 4 Pickup split bills omitted by the same report.</li><li>Total split tender omitted from that export: <b>AED 1,181.90</b>. It is not unpaid revenue.</li><li>The correct bill/payment ledger already contains these amounts and reconciles to <b>AED 108,172.75</b>.</li><li><b>AED 117.00</b> of Dine In tendered as Talabat and <b>AED 65.60</b> of Talabat tendered as Keeta remain genuine channel-mapping reviews.</li><li>Payment allocation must never determine output VAT filing; the invoice supply ledger is the VAT control.</li></ul></Info>
        <section className="resolutionPanel"><div className="panelHead"><div><h3>Exact resolution for AED 1,181.90</h3><p>The error is in the Tax Submission report’s payment extraction—not in invoice sales.</p></div><a className="downloadBtn" href="/downloads/TMBill_Reconciliation_Control.xlsx" download>Download reconciliation Excel ↓</a></div>
          <ol><li><b>Fix the Tax Submission report query/service.</b> It currently captures only one component of Card/Cash split bills. Join all active payment-allocation records by Order ID.</li><li><b>Use the correct grain.</b> One row must represent one invoice × tender allocation, including allocation ID, amount, status, reference and timestamp.</li><li><b>Add an invoice control.</b> For every fulfilled invoice: tender allocations + wallet + due = invoice total, adjusted for valid refunds/credit notes.</li><li><b>Backfill the 15 split invoices.</b> Mark them “Split confirmed – correct in bill ledger / incomplete in Tax Submission report,” not unpaid or missing.</li><li><b>Keep the 3 channel mismatches separate.</b> They require mapping review but have no monetary shortfall.</li><li><b>Rebuild the reports.</b> The payment matrix must total AED 108,172.75. VAT filing remains AED 5,124.92 from the invoice VAT ledger, not from payments.</li></ol>
        </section>
        {paymentRecon&&<section className="panel"><div className="panelHead"><div><h3>Affected invoices requiring payment-report correction</h3><p>{paymentRecon.summary.exceptionInvoices} exceptions: 15 missing allocations plus 3 cross-channel validations.</p></div></div><DataGrid id="payment-exceptions" rows={paymentRecon.exceptions} onRowClick={x=>{setSelectedBill(x.billNo);setTab("bills")}} columns={[
          {key:"billNo",label:"Bill",render:x=><b>#{x.billNo}</b>},{key:"orderId",label:"Order ID"},{key:"date",label:"Date"},{key:"orderType",label:"Order type"},{key:"paymentMode",label:"Payment mode"},{key:"correctTotal",label:"Correct bill tender",numeric:true,render:x=>money(x.correctTotal)},{key:"tmbillTotal",label:"Tax report captured",numeric:true,render:x=>money(x.tmbillTotal)},{key:"missing",label:"Secondary split omitted",numeric:true,render:x=>x.missing?money(x.missing):"—"},{key:"delta_card",label:"Omitted Card",numeric:true,render:x=>x.delta_card?money(x.delta_card):"—"},{key:"delta_cash",label:"Omitted Cash",numeric:true,render:x=>x.delta_cash?money(x.delta_cash):"—"},{key:"issueType",label:"Issue type"},{key:"reconciliationStatus",label:"Resolution",render:x=><Badge tone={x.channelMismatch?"bad":"good"}>{x.reconciliationStatus}</Badge>}
        ]} totals={{billNo:"CONTROL",missing:money(paymentRecon.summary.missing),reconciliationStatus:`${paymentRecon.summary.splitPaymentsConfirmed} split confirmed · ${paymentRecon.summary.channelMismatches} mapping reviews`}}/></section>}
        <Info title="Why the charge gap is AED 61.00" tone="amber"><ul><li>The invoice-level Sales Report records <b>AED 581.00</b> under Total Charges.</li><li>The DSR Month-wise and Bill-wise reports expose only <b>AED 520.00</b> as Delivery/Total Charges.</li><li>The unresolved difference is therefore <b>AED 61.00</b>.</li><li>The Sales Report contains twelve columns all named “Delivery Charges,” so the missing AED 61 cannot be mapped to a named charge type.</li><li>Fix: replace those duplicate columns with a normalized invoice-charge table containing charge ID, name, net, VAT rate, VAT and gross. Then require charge rows to reconcile exactly to invoice Total Charges.</li></ul></Info>
        <section className="sourceAudit"><div className="panelHead"><div><h3>Every supplied workbook assessed</h3><p>Select a report for its confirmed issue, correction and original header scope.</p></div></div><div className="auditLayout"><div className="auditList">{sourceAudits.map(a=><button key={a.name} className={selectedSource===a.name?"active":""} onClick={()=>setSelectedSource(a.name)}><Badge tone={a.status==="pass"?"good":a.status==="warn"?"warn":"bad"}>{a.status}</Badge><span>{a.name}</span></button>)}</div>{(()=>{const a=sourceAudits.find(x=>x.name===selectedSource)!;return <article className="auditDetail"><span>{a.file}</span><h2>{a.name}</h2><Badge tone={a.status==="pass"?"good":a.status==="warn"?"warn":"bad"}>{a.status.toUpperCase()}</Badge><h4>What TMBill gets wrong or omits</h4><p>{a.issue}</p><h4>Required correction</h4><p>{a.fix}</p><h4>Headers reviewed</h4><div className="headerCloud">{a.headers.split(", ").map(h=><Badge key={h}>{h}</Badge>)}</div></article>})()}</div></section>
      </>}

      {tab==="separate"&&separateAudit&&<>
        <section className="uploadAudit"><div><Badge tone="good">REUSABLE SALES AUDITOR</Badge><h2>Audit any TMBill Sales Report</h2><p>Upload a compatible Sales Report Excel file. It is analysed locally in your browser and never mixed with the June control dataset.</p></div><label className="uploadDrop"><input type="file" accept=".xlsx,.xls" onChange={e=>handleSalesUpload(e.target.files?.[0])}/><b>Choose Sales Report Excel</b><span>Required: Order ID, Bill, Subtotal, Discount, Tax, Charges, Total and Status.</span></label>{uploadError&&<div className="uploadError">{uploadError}</div>}</section>
        <section className="hero slim isolatedHero"><div><Badge tone="warn">ISOLATED DATASET · NOT COMBINED WITH JUNE</Badge><h2>Sales Report audit: {separateAudit.summary.periodFrom} – {separateAudit.summary.periodTo}</h2><p>Source: {separateAudit.summary.file}. Uploading another workbook replaces only this tab’s working dataset.</p></div><div className="heroScore"><b>{money(separateAudit.summary.total)}</b><span>fulfilled invoice revenue</span><small>{separateAudit.summary.fulfilled} bills · #{separateAudit.summary.billFrom}–#{separateAudit.summary.billTo}</small></div></section>
        <section className="kpiGrid">
          <Kpi label="Subtotal before discount" value={money(separateAudit.summary.subtotal)} sub="Report invoice subtotal · reconciled"/>
          <Kpi label="Total discounts" value={money(separateAudit.summary.discount)} sub="Discount components reconcile exactly"/>
          <Kpi label="Reported VAT" value={money(separateAudit.summary.vat)} sub={`Invoice aggregate control ${money(separateAudit.summary.vatControl)}`}/>
          <Kpi label="Charges" value={money(separateAudit.summary.charges)} sub="13 charged invoices · report footer shows only AED 75" tone="bad"/>
          <Kpi label="Net revenue ex VAT" value={money(separateAudit.summary.netExVat)} sub="Total revenue less reported VAT"/>
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
          {key:"orderType",label:"Order type",render:x=><b>{x.orderType}</b>},{key:"orders",label:"Orders",numeric:true},
          {key:"card",label:"Card",numeric:true,render:x=><TenderValue row={x} tender="card"/>},{key:"cash",label:"Cash",numeric:true,render:x=><TenderValue row={x} tender="cash"/>},
          {key:"talabat",label:"Talabat",numeric:true,render:x=><TenderValue row={x} tender="talabat"/>},{key:"deliveroo",label:"Deliveroo",numeric:true,render:x=><TenderValue row={x} tender="deliveroo"/>},{key:"keeta",label:"Keeta",numeric:true,render:x=><TenderValue row={x} tender="keeta"/>},
          {key:"correctSales",label:"Correct sales",numeric:true,render:x=>money(x.correctSales)},{key:"invoiceVat",label:"Invoice VAT",numeric:true,render:x=>money(x.invoiceVat)},{key:"taxReportCaptured",label:"Tax report captured",numeric:true,render:x=>money(x.taxReportCaptured)}
        ]} totals={{orderType:"TOTAL",orders:num(separateAudit.summary.fulfilled),card:money(separateMatrix.reduce((a,x)=>a+x.card,0)),cash:money(separateMatrix.reduce((a,x)=>a+x.cash,0)),talabat:money(separateMatrix.reduce((a,x)=>a+x.talabat,0)),deliveroo:money(separateMatrix.reduce((a,x)=>a+x.deliveroo,0)),keeta:money(separateMatrix.reduce((a,x)=>a+x.keeta,0)),correctSales:money(separateAudit.summary.total),invoiceVat:money(separateAudit.summary.vat),taxReportCaptured:money(separateMatrix.reduce((a,x)=>a+x.taxReportCaptured,0))}}/></section>
        <div className="filters"><label className="search"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search bill, order ID or order type…"/></label><select value={separateStatus} onChange={e=>setSeparateStatus(e.target.value)}><option value="All">All assessments</option><option value="pass">Pass</option><option value="review">Review</option></select><select value={separateIssue} onChange={e=>setSeparateIssue(e.target.value)}><option value="All">All issues</option><option>Charges present with zero Tax On Charges</option><option>Charge components do not equal Total Charges</option><option>Split payment amounts unavailable</option><option>VAT differs from aggregate 5/105 control</option></select><button className="downloadBtn" onClick={downloadSeparateAudit}>Download current audit ↓</button><span className="resultCount">{num(separateRows.length)} bills</span></div>
        <DataGrid id="separate-invoices" rows={separateRows} selectedKey={selectedSeparateBill} onRowClick={x=>{setSelectedSeparateBill(x.billNo);setTimeout(()=>document.getElementById("separate-bill-preview")?.scrollIntoView({behavior:"smooth",block:"start"}),60)}} columns={[
          {key:"billNo",label:"Bill",render:x=><b>#{x.billNo}</b>},{key:"id",label:"Order ID"},{key:"date",label:"Date"},{key:"orderType",label:"Order type"},{key:"paymentMode",label:"Payment label"},{key:"subtotal",label:"Subtotal",numeric:true,render:x=>money(x.subtotal)},{key:"discount",label:"Discount",numeric:true,render:x=>money(x.discount)},{key:"vat",label:"Reported VAT",numeric:true,render:x=>money(x.vat)},{key:"vatControl",label:"VAT control",numeric:true,render:x=>money(x.vatControl)},{key:"vatVariance",label:"VAT variance",numeric:true,render:x=>x.vatVariance?money(x.vatVariance):"—"},{key:"charges",label:"Total charges",numeric:true,render:x=>x.charges?money(x.charges):"—"},{key:"chargeParts",label:"Visible charge component",numeric:true,render:x=>x.chargeParts?money(x.chargeParts):"—"},{key:"taxOnCharges",label:"Tax on charges",numeric:true,render:x=>x.taxOnCharges?money(x.taxOnCharges):"—"},{key:"total",label:"Total",numeric:true,render:x=>money(x.total)},{key:"issues",label:"Issues",render:x=>x.issues.length?<span className="issueText">{x.issues.join(" · ")}</span>:"—"},{key:"assessment",label:"Assessment",render:x=><Badge tone={x.assessment==="pass"?"good":"warn"}>{x.assessment}</Badge>}
        ]} totals={{billNo:"TOTAL",subtotal:money(separateRows.reduce((a,x)=>a+x.subtotal,0)),discount:money(separateRows.reduce((a,x)=>a+x.discount,0)),vat:money(separateRows.reduce((a,x)=>a+x.vat,0)),charges:money(separateRows.reduce((a,x)=>a+x.charges,0)),total:money(separateRows.reduce((a,x)=>a+x.total,0))}}/>
        {separateSelected&&<section className="separatePreview" id="separate-bill-preview"><div className="previewTop"><div><span>INVOICE PREVIEW</span><h2>Bill #{separateSelected.billNo}</h2><p>{separateSelected.date} · {separateSelected.time} · {separateSelected.orderType}</p></div><div className="previewLinks"><a href={`https://backoffice.tmbill.com/ebill/${separateSelected.id}`} target="_blank" rel="noreferrer">Open official TMBill eBill ↗</a><button onClick={()=>setSelectedSeparateBill("")}>Close</button></div></div><div className="receiptFrame local"><div className="receiptLocal"><div className="receiptAccent"/><div className="receiptBrand">MORBIDO</div><h3>Morbido Express Restaurant</h3><p>TAX Invoice · Audit preview</p><div className="receiptRule"/><b className="orderPill">{separateSelected.orderType}</b><h2>Order No: {separateSelected.billNo}</h2><small>Order ID: {separateSelected.id}</small><div className="receiptSection"><h5>Order Details</h5><div><span>Status</span><b>{separateSelected.status}</b></div><div><span>Table</span><b>{separateSelected.table||"—"}</b></div><div><span>User</span><b>{separateSelected.user}</b></div><div><span>Payment label</span><b>{separateSelected.paymentMode}</b></div><div><span>Guests</span><b>{separateSelected.guests||"—"}</b></div></div><div className="receiptSection summary"><div><span>Subtotal</span><b>{money(separateSelected.subtotal)}</b></div><div><span>Discount</span><b>{money(separateSelected.discount)}</b></div><div><span>Charges</span><b>{money(separateSelected.charges)}</b></div><div><span>Net without VAT</span><b>{money(separateSelected.total-separateSelected.vat)}</b></div><div><span>VAT</span><b>{money(separateSelected.vat)}</b></div><div className="grand"><span>Grand Total</span><b>{money(separateSelected.total)}</b></div></div><p className="thanks">Audit reconstruction</p><small>Item lines and split-tender amounts are not present in this Sales Report.</small></div></div><Info title="Invoice assessment" tone={separateSelected.issues.length?"amber":"blue"}>{separateSelected.issues.length?separateSelected.issues.join(" · "):"This invoice passes the controls available in the uploaded Sales Report."}</Info></section>}
      </>}

      {tab==="zreport"&&<>
        <section className="zSheet"><div className="zHead"><div><span>TAX / MANAGEMENT CONTROL REPORT</span><h2>Morbido Express Restaurant</h2><p>The SS · 01 Jun 2026 12:45 PM – 30 Jun 2026 11:58 PM</p></div><Badge tone="good">CLOSED · RECONCILED</Badge></div>
          <div className="zMeta"><div><span>First / last bill</span><b>#2313 – #4141</b></div><div><span>Fulfilled bills</span><b>1,816</b></div><div><span>Cancelled bills</span><b>13 · {money(1094.50)}</b></div><div><span>Guests</span><b>1,817</b></div></div>
          <section className="zBlock"><h3>Sales and VAT bridge</h3>{[
            ["Original item gross",118066],["Less: item discounts",-13],["Less: order discounts",-10461.25],["Taxable item sales incl VAT",107591.75],["Net taxable sales ex VAT",102466.82],["Output VAT – canonical lines",5124.93],["TMBill reported output VAT",5124.92],["VAT rounding control",-0.01],["Charges collected",581],["TOTAL REVENUE",108172.75]
          ].map(([n,v])=><div className={n==="TOTAL REVENUE"?"grand":""} key={n as string}><span>{n}</span><b>{money(v as number)}</b></div>)}</section>
          <div className="zColumns"><section className="zBlock"><h3>Order types</h3>{paymentMatrix.map(x=><div key={x.orderType}><span>{x.orderType} ({x.orders})</span><b>{money(x.sales)}</b></div>)}<div className="grand"><span>Total</span><b>{money(108172.75)}</b></div></section>
          <section className="zBlock"><h3>Payment tenders</h3>{paymentTotals.map(([n,v,c])=><div key={n}><span>{n} ({c})</span><b>{money(v)}</b></div>)}<div className="grand"><span>Total</span><b>{money(108172.75)}</b></div></section></div>
          <section className="zBlock"><h3>Required review controls</h3><div><span>Confirmed item/order conflict</span><b>#2990 · corrected</b></div><div><span>Discount allocation rounding</span><b>6 invoices</b></div><div><span>VAT line rounding</span><b>252 invoices · max AED 0.02</b></div><div><span>Tax-payment allocation shortfall</span><b>{money(1181.90)}</b></div><div><span>Charges without charge VAT</span><b>{money(581)}</b></div></section>
          <Info title="Z-report sign-off" tone="amber">Do not use “payment collected” as the sales or VAT basis. Sign off the invoice supply ledger, discount bridge, VAT ledger, charge-tax classification and payment reconciliation separately. The AED 581 charge treatment remains unresolved until charge types/contracts are reviewed.</Info>
        </section>
      </>}

      {tab==="tax"&&<>
        <section className="hero slim"><div><Badge tone="good">No material invoice VAT exception found</Badge><h2>VAT total is plausible; the audit trail is not sufficient.</h2><p>Every closed invoice is tested against an inclusive 5% aggregate control. Differences are limited to AED ±0.02, consistent with line rounding. The item export still cannot reproduce the tax because its post-discount tax bases are unreliable.</p></div><div className="heroScore"><b>{money(5124.92)}</b><span>TMBill invoice VAT</span><small>Output VAT before credit-note adjustments</small></div></section>
        <section className="kpiGrid">
          <Kpi label="Taxable gross incl VAT" value={money(107591.75)} sub="Subtotal AED 118,053 − discounts AED 10,461.25"/>
          <Kpi label="TMBill reported VAT" value={money(5124.92)} sub="Sum of VAT on 1,816 supplied invoices"/>
          <Kpi label="Invoice aggregate control" value={money(5124.03)} sub="Σ round(invoice taxable gross × 5/105, 2)" tone="warn"/>
          <Kpi label="Single period control" value={money(5123.42)} sub="Round(period taxable gross × 5/105, 2)" tone="warn"/>
          <Kpi label="Exact invoice controls" value="1,564" sub="Reported VAT equals invoice aggregate control"/>
          <Kpi label="Rounding differences" value="252" sub="All differences are AED 0.01 or AED 0.02" tone="warn"/>
          <Kpi label="Material exceptions" value="0" sub="No invoice differs by more than AED 0.02"/>
          <Kpi label="Invoice vs period rounding" value={money(1.50)} sub="Not automatically an under/overpayment" tone="warn"/>
        </section>
        <Info title="Critical compliance correction" tone="red">VAT payable is generally based on taxable supplies under the UAE date-of-supply rules—not only cash collected. An unpaid supplied invoice can still create output VAT. For this dataset, due receivables are zero, but the calculation engine must never equate “VAT payable” with “VAT collected in cash”.</Info>
        <Info title="Highest unresolved VAT risk" tone="amber">AED 581.00 of delivery/other charges is included in customer totals while “Tax on Charges” is AED 0.00. If these charges are consideration for, or ancillary to, the standard-rated restaurant supply, they may also require VAT. If treated as VAT-inclusive, the potential VAT is AED 27.67. TMBill must store each charge type, tax treatment, rate and VAT amount; a UAE tax adviser should confirm the classification before filing.</Info>
        <Info title="Two decimals or three?" tone="blue">Store calculations internally at high precision (at least 4–6 decimal places), but round the VAT amount payable on each tax invoice to the nearest fils—two decimal places—using mathematical half-up rounding. Three-decimal VAT can be retained only as an internal calculation trace; it should not be the posted invoice or filing amount. The system must document whether it rounds per line or at invoice total and use that method consistently.</Info>
        <Info title="Are service and delivery charges taxable?" tone="red">Generally, a compulsory service, delivery, packaging or similar fee imposed by the restaurant as part of making the taxable food supply forms part of consideration and follows the supply’s VAT treatment. A true disbursement paid in the customer’s name and account can differ. Because TMBill provides twelve anonymous “Delivery Charges” columns and zero Tax On Charges, the AED 581 cannot be safely classified. The immediate fix is a charge ledger with charge name, compulsory/optional flag, principal/agent treatment, VAT rate, net, VAT and gross.</Info>
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
        <section className="hero slim"><div><Badge tone="good">Corrected canonical category view</Badge><h2>Every category now uses actual transaction prices and allocated discounts.</h2><p>Original line gross is reduced by item discount and eligible allocated order discount, then divided into VAT and net revenue.</p></div><div className="heroScore"><b>{money(+s.taxableGrossInclVat)}</b><span>corrected item sales incl VAT</span><small>Charges shown separately: {money(+s.charges)}</small></div></section>
        <DataGrid id="category-ledger" rows={data.categories} columns={[
          {key:"name",label:"Category"},{key:"qty",label:"Qty",numeric:true},{key:"gross",label:"Original gross",numeric:true,render:x=>money(x.gross)},{key:"itemDiscount",label:"Item discount",numeric:true,render:x=>money(x.itemDiscount)},{key:"orderDiscount",label:"Order discount",numeric:true,render:x=>money(x.orderDiscount)},{key:"discount",label:"Total discount",numeric:true,render:x=>money(x.discount)},{key:"net",label:"Sold incl VAT",numeric:true,render:x=>money(x.net)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"netExVat",label:"Sold ex VAT",numeric:true,render:x=>money(x.netExVat)},{key:"mix",label:"Mix %",numeric:true,value:x=>x.net/+s.taxableGrossInclVat*100,render:x=>`${(x.net/+s.taxableGrossInclVat*100).toFixed(2)}%`}
        ]} totals={{name:"TOTAL",qty:num(data.categories.reduce((a,c)=>a+c.qty,0)),gross:money(data.categories.reduce((a,c)=>a+c.gross,0)),itemDiscount:money(data.categories.reduce((a,c)=>a+c.itemDiscount,0)),orderDiscount:money(data.categories.reduce((a,c)=>a+c.orderDiscount,0)),discount:money(data.categories.reduce((a,c)=>a+c.discount,0)),net:money(data.categories.reduce((a,c)=>a+c.net,0)),vat:money(data.categories.reduce((a,c)=>a+c.vat,0)),netExVat:money(data.categories.reduce((a,c)=>a+c.netExVat,0)),mix:"100.00%"}}/>
        <Info title="Reconciliation rule" tone="blue">Corrected category sales including VAT reconcile to AED 107,591.75. Add AED 581.00 separately classified charges to reach AED 108,172.75 invoice revenue. Charges must not be forced into categories without an explicit allocation rule.</Info>
      </>}

      {tab==="reports"&&<>
        <section className="hero slim"><div><Badge tone="good">20 governed views</Badge><h2>Ideal restaurant reporting library</h2><p>Each view below uses the same invoice, item, payment and event facts. Select a report to understand its purpose and grain.</p></div></section>
        <div className="filters"><label className="search"><span>⌕</span><input value={reportSearch} onChange={e=>setReportSearch(e.target.value)} placeholder="Search reports and definitions…"/></label><span className="resultCount">{reportSpecs.filter(x=>x.join(" ").toLowerCase().includes(reportSearch.toLowerCase())).length} reports</span></div>
        <section className="reportGrid">{reportSpecs.filter(x=>x.join(" ").toLowerCase().includes(reportSearch.toLowerCase())).map((r,i)=><button className={`reportCard ${selectedReport===r[0]?"active":""}`} onClick={()=>{setSelectedReport(r[0]);setTimeout(()=>document.getElementById("ideal-report-preview")?.scrollIntoView({behavior:"smooth",block:"start"}),80)}} key={r[0]}><span>{String(i+1).padStart(2,"0")}</span><div><h3>{r[0]}</h3><Badge>{r[1]}</Badge><p>{r[2]}</p><strong>Preview report ↓</strong></div></button>)}</section>
        <section className="reportPreview" id="ideal-report-preview"><div className="previewTop"><div><span>REPORT PREVIEW</span><h2>{selectedReport}</h2><p>{reportSpecs.find(x=>x[0]===selectedReport)?.[2]}</p></div><Badge tone="good">Canonical design</Badge></div>
          <div className="previewControls"><div><small>OUTLET</small><b>Morbido Express Restaurant</b></div><div><small>PERIOD</small><b>01–30 Jun 2026</b></div><div><small>STATUS</small><b>Supplied / fulfilled</b></div><div><small>BASIS</small><b>VAT inclusive + explicit VAT</b></div></div>
          {selectedReport==="Sales / Z Summary"?<DataGrid id="report-z" rows={[{period:"01–30 Jun 2026",bills:+s.fulfilledInvoices,gross:+s.correctedGrossBeforeDiscount,itemDiscount:+s.canonicalItemDiscount,orderDiscount:+s.canonicalOrderDiscount,charges:+s.charges,vat:+s.vat,netExVat:+s.netSales,total:+s.grossSales}]} columns={[
            {key:"period",label:"Period"},{key:"bills",label:"Bills",numeric:true},{key:"gross",label:"Original gross",numeric:true,render:x=>money(x.gross)},{key:"itemDiscount",label:"Item discount",numeric:true,render:x=>money(x.itemDiscount)},{key:"orderDiscount",label:"Order discount",numeric:true,render:x=>money(x.orderDiscount)},{key:"charges",label:"Charges",numeric:true,render:x=>money(x.charges)},{key:"vat",label:"VAT",numeric:true,render:x=>money(x.vat)},{key:"netExVat",label:"Net ex VAT",numeric:true,render:x=>money(x.netExVat)},{key:"total",label:"Total revenue",numeric:true,render:x=>money(x.total)}
          ]}/>
          :selectedReport==="Payment Type Summary"?<DataGrid id="report-payments" rows={paymentTotals.map(([name,total,allocations])=>({name,total,allocations,mix:total/+s.grossSales*100}))} columns={[{key:"name",label:"Tender"},{key:"allocations",label:"Allocations",numeric:true},{key:"total",label:"Collected",numeric:true,render:x=>money(x.total)},{key:"mix",label:"Mix %",numeric:true,render:x=>`${x.mix.toFixed(2)}%`}]} totals={{name:"TOTAL",allocations:"1,831",total:money(+s.grossSales),mix:"100.00%"}}/>
          :selectedReport==="Discount Summary"?<><DataGrid id="report-discounts" rows={[{level:"Confirmed item adjustment",bills:1,discount:+s.canonicalItemDiscount,status:"Minimum confirmed"},{level:"Invoice/order discounts",bills:587,discount:+s.canonicalOrderDiscount,status:"Reconciled"}]} columns={[{key:"level",label:"Discount level"},{key:"bills",label:"Bills",numeric:true},{key:"discount",label:"Discount",numeric:true,render:x=>money(x.discount)},{key:"status",label:"Control"}]} totals={{level:"CONFIRMED MINIMUM",discount:money(+s.canonicalTotalDiscount)}}/><Info title="Completeness limitation" tone="red">The June item exports confirm one normally chargeable line sold at zero: Fattoush on bill #2990. They do not retain original transaction-time menu price or a separate item-discount field, so partial item-level discounts cannot be certified. Current-menu comparisons produce hundreds of price variances that may reflect menu-version changes rather than discounts.</Info></>
          :selectedReport==="Bill Summary"?<DataGrid id="report-bills" rows={data.invoices} columns={[
            {key:"billNo",label:"Bill"},{key:"id",label:"Order ID"},{key:"date",label:"Supply date"},{key:"orderType",label:"Order type"},
            {key:"canonicalGrossBeforeDiscount",label:"Gross",numeric:true,render:x=>money(x.canonicalGrossBeforeDiscount)},
            {key:"canonicalItemDiscount",label:"Item discount",numeric:true,render:x=>money(x.canonicalItemDiscount)},
            {key:"canonicalOrderDiscount",label:"Order discount",numeric:true,render:x=>money(x.canonicalOrderDiscount)},
            {key:"canonicalVat",label:"VAT",numeric:true,render:x=>money(x.canonicalVat)},{key:"charges",label:"Charges",numeric:true,render:x=>money(x.charges)},{key:"total",label:"Total",numeric:true,render:x=>money(x.total)}
          ]} totals={{billNo:"TOTAL",canonicalGrossBeforeDiscount:money(+s.correctedGrossBeforeDiscount),canonicalItemDiscount:money(+s.canonicalItemDiscount),canonicalOrderDiscount:money(+s.canonicalOrderDiscount),canonicalVat:money(5124.93),charges:money(+s.charges),total:money(+s.grossSales)}}/>
          :selectedReport==="Sold Items Summary"?<><DataGrid id="report-items" rows={data.topItems} columns={[
            {key:"name",label:"Item"},{key:"qty",label:"Qty",numeric:true},{key:"gross",label:"Original gross",numeric:true,render:x=>money(x.gross)},{key:"discount",label:"Total discount",numeric:true,render:x=>money(x.discount)},{key:"net",label:"Net incl VAT",numeric:true,render:x=>money(x.net)}
          ]} totals={{name:"ITEM SALES TOTAL",qty:num(data.topItems.reduce((a,x)=>a+x.qty,0)),gross:money(data.topItems.reduce((a,x)=>a+x.gross,0)),discount:money(data.topItems.reduce((a,x)=>a+x.discount,0)),net:money(data.topItems.reduce((a,x)=>a+x.net,0))}}/><SalesReconciliation itemSales={data.topItems.reduce((a,x)=>a+x.net,0)} charges={+s.charges} actualSales={+s.grossSales}/></>
          :selectedReport==="Category Summary"||selectedReport==="Product Group Summary"?<><DataGrid id="report-category" rows={data.categories} columns={[
            {key:"name",label:"Category"},{key:"qty",label:"Qty",numeric:true},{key:"gross",label:"Original gross",numeric:true,render:x=>money(x.gross)},{key:"discount",label:"Allocated discounts",numeric:true,render:x=>money(x.discount)},{key:"net",label:"Net incl VAT",numeric:true,render:x=>money(x.net)}
          ]} totals={{name:"ITEM SALES TOTAL",qty:num(data.categories.reduce((a,x)=>a+x.qty,0)),gross:money(data.categories.reduce((a,x)=>a+x.gross,0)),discount:money(data.categories.reduce((a,x)=>a+x.discount,0)),net:money(data.categories.reduce((a,x)=>a+x.net,0))}}/><SalesReconciliation itemSales={data.categories.reduce((a,x)=>a+x.net,0)} charges={+s.charges} actualSales={+s.grossSales}/></>
          :selectedReport==="Cancel Items Summary"&&kotAudit?<DataGrid id="report-cancel-items" rows={kotAudit.cancellations} columns={[
            {key:"punchTime",label:"KOT punch time"},{key:"billNo",label:"Bill"},{key:"orderId",label:"Order ID"},{key:"kotId",label:"KOT ID"},{key:"item",label:"Item"},{key:"qty",label:"Qty",numeric:true},{key:"listedValue",label:"Original list value",numeric:true,render:x=>money(x.listedValue)},{key:"user",label:"User"},{key:"reason",label:"Reason",render:x=>x.reason||<Badge tone="bad">Missing</Badge>},{key:"status",label:"Status",render:x=><Badge tone="bad">{x.status}</Badge>}
          ]} totals={{billNo:"TOTAL",qty:num(kotAudit.summary.cancelledQuantity),listedValue:money(kotAudit.summary.cancelledListValue)}}/>
          :selectedReport==="Order Type Summary"?<DataGrid id="report-orders" rows={orderTotals.map(([name,total,bills])=>({name,total,bills,vat:total*5/105,mix:total/+s.grossSales*100}))} columns={[{key:"name",label:"Order type"},{key:"bills",label:"Bills",numeric:true},{key:"total",label:"Sales",numeric:true,render:x=>money(x.total)},{key:"vat",label:"VAT control",numeric:true,render:x=>money(x.vat)},{key:"mix",label:"Mix %",numeric:true,render:x=>`${x.mix.toFixed(2)}%`}]} totals={{name:"TOTAL",bills:num(+s.fulfilledInvoices),total:money(+s.grossSales),vat:money(+s.vat),mix:"100.00%"}}/>
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
          <article className="guideCard"><span>01</span><h3>Start at gross sales</h3><p>Fulfilled supplied invoices total AED 108,172.75. This is the master financial control for the period.</p></article>
          <article className="guideCard"><span>02</span><h3>Bridge VAT explicitly</h3><p>Gross sales − invoice VAT AED 5,124.92 = net sales ex VAT AED 103,047.83.</p></article>
          <article className="guideCard"><span>03</span><h3>Trace dimensions to lines</h3><p>Category, kitchen and item reports must use validated item lines with order discounts allocated.</p></article>
          <article className="guideCard"><span>04</span><h3>Keep cancellations separate</h3><p>13 cancelled orders, 103 cancelled KOT items and post-sale refunds are different events.</p></article>
          <article className="guideCard"><span>05</span><h3>Explode split tenders</h3><p>1,831 payment allocations across 1,816 invoices is valid. Missing allocation amounts are not.</p></article>
          <article className="guideCard"><span>06</span><h3>Version menu prices</h3><p>Current-price differences become provable only when item IDs and effective-dated price versions are stored.</p></article>
        </section>
        <section className="panel"><div className="panelHead"><div><h3>Canonical equations</h3><p>These equations should appear in dashboard tooltips and report documentation.</p></div></div>
          <div className="formulaList">
            <div><code>Gross item value</code><b>Σ(quantity × transaction unit price)</b></div>
            <div><code>Post-discount line</code><b>gross − item discount − allocated order discount</b></div>
            <div><code>Invoice VAT</code><b>Σ line VAT + charge VAT</b></div>
            <div><code>Gross sales</code><b>net items incl VAT + charges incl VAT + round-off</b></div>
            <div><code>Net sales ex VAT</code><b>gross sales − total VAT</b></div>
            <div><code>Open due</code><b>gross sales − payments − wallet − credits</b></div>
          </div>
        </section>
        <Info title="Inventory warning" tone="red">Raw-material cost is zero because recipes and costs are not configured. AED 118,053 is gross item subtotal—not gross margin. A 100% gross-profit percentage must not be displayed until COGS exists.</Info>
      </>}
    </main>
  </div>
}
