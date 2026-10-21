import ExcelJS from "exceljs";
import { apiClient } from "../api/client";

export type TransferStatus = "Draft" | "Pending Receipt" | "Returned for Correction" | "Completed";
export type TransferSection = "Criminal" | "Civil" | "ADMIN 1" | "ADMIN 2" | "ADMIN 3";

export interface TransferAttorney { user_id: number; full_name: string; }
export interface TransferCase {
  case_id: number;
  case_type: "Criminal" | "Civil";
  control_no: string;
  title_of_case: string;
  case_no: string;
  date_received: string;
  status_of_case: string;
  cause_of_action: string;
  assigned_pao: string;
}
export interface TransferAdminEntry {
  entry_id?: number;
  section: "ADMIN 1" | "ADMIN 2" | "ADMIN 3";
  row_number: number;
  control_no: string;
  title_of_case: string;
  docket_no: string;
  date_received: string;
  status_of_case: string;
  cause_of_action: string;
}
export interface TransferItem { item_id: number; case_id: number; section: "Criminal" | "Civil"; snapshot: TransferCase; }
export interface CaseTransfer {
  report_id: number;
  title: string;
  report_date: string;
  date_received: string | null;
  status: TransferStatus;
  remarks: string;
  version: number;
  from_user_id: number;
  from_user_name: string;
  to_user_id: number;
  to_user_name: string;
  submitted_at: string | null;
  returned_at: string | null;
  confirmed_at: string | null;
  created_at: string;
  updated_at: string;
  items: TransferItem[];
  administrative_entries: TransferAdminEntry[];
}
export interface TransferPayload {
  title: string;
  report_date: string;
  to_user_id: number;
  case_ids: number[];
  administrative_entries: TransferAdminEntry[];
  remarks?: string;
}

export async function listTransferAttorneys() {
  return (await apiClient.get<TransferAttorney[]>("/case-transfers/attorneys")).data;
}
export async function listTransferCases(caseType: "Criminal" | "Civil", search = "") {
  return (await apiClient.get<TransferCase[]>("/case-transfers/cases", { params: { case_type: caseType, search: search || undefined } })).data;
}
export async function listCaseTransfers() {
  return (await apiClient.get<CaseTransfer[]>("/case-transfers")).data;
}
export async function getCaseTransfer(id: number) {
  return (await apiClient.get<CaseTransfer>(`/case-transfers/${id}`)).data;
}
export async function createCaseTransfer(payload: TransferPayload) {
  return (await apiClient.post<CaseTransfer>("/case-transfers", payload)).data;
}
export async function updateCaseTransfer(id: number, payload: TransferPayload) {
  return (await apiClient.patch<CaseTransfer>(`/case-transfers/${id}`, payload)).data;
}
export async function submitCaseTransfer(id: number) {
  return (await apiClient.post<CaseTransfer>(`/case-transfers/${id}/submit`)).data;
}
export async function returnCaseTransfer(id: number, remarks: string) {
  return (await apiClient.post<CaseTransfer>(`/case-transfers/${id}/return`, { remarks })).data;
}
export async function confirmCaseTransfer(id: number) {
  return (await apiClient.post<CaseTransfer>(`/case-transfers/${id}/confirm`)).data;
}

function formatTransferDate(value?: string | null) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleDateString("en-PH", { year: "numeric", month: "short", day: "numeric" });
}

async function loadImage(url: string) {
  const response = await fetch(url);
  if (!response.ok) throw new Error("Unable to load transfer report logo.");
  return response.arrayBuffer();
}

export async function buildCaseTransferExcel(report: CaseTransfer, logo1Url: string, logo2Url: string) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "JurisGuard";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("Transfer of Cases", { views: [{ showGridLines: false }] });
  const headings = ["NO.", "CONTROL NO.", "TITLE OF THE CASE", "DOCKET NO.", "DATE RECEIVED", "STATUS OF THE CASE", "CAUSE OF ACTION"];
  sheet.columns = [
    { key: "no", width: 5 }, { key: "control", width: 14 }, { key: "title", width: 25 },
    { key: "docket", width: 12 }, { key: "date", width: 13 }, { key: "status", width: 15 }, { key: "cause", width: 26 },
  ];
  sheet.pageSetup = { paperSize: 9, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 }, horizontalDpi: 300, verticalDpi: 300 };
  const border = { style: "thin" as const, color: { argb: "FF000000" } };
  const allBorders = { top: border, left: border, bottom: border, right: border };
  const center = { horizontal: "center" as const, vertical: "middle" as const, wrapText: true };
  const left = { horizontal: "left" as const, vertical: "top" as const, wrapText: true };
  const blackFont = { name: "Arial", size: 9, color: { argb: "FF000000" } };

  ["A1:G1", "A2:G2", "A3:G3", "A4:G4", "A5:G5", "A7:G7", "A8:G8"].forEach((range) => sheet.mergeCells(range));
  const headingRows: Array<[number, string, number, boolean]> = [
    [1, "Republic of the Philippines", 10, false], [2, "Department of Justice", 10, false], [3, "PUBLIC ATTORNEY'S OFFICE", 12, true],
    [4, "Regional Office No. XI", 10, false], [5, "District Office: Panabo City", 10, false], [7, "TRANSFER OF CASES", 14, true], [8, `As of ${formatTransferDate(report.report_date)}`, 10, false],
  ];
  headingRows.forEach(([number, value, size, bold]) => { const cell = sheet.getCell(`A${number}`); cell.value = value; cell.font = { ...blackFont, size, bold }; cell.alignment = center; sheet.getRow(number).height = number === 7 ? 22 : 16; });
  sheet.getRow(6).height = 6;
  const details = [["A10", `FROM: ${report.from_user_name}`], ["A11", `Date Transferred: ${formatTransferDate(report.submitted_at || report.report_date)}`], ["A12", "Signature: ______________________________"], ["E10", `TO: ${report.to_user_name}`], ["E11", `Date Received: ${formatTransferDate(report.date_received)}`], ["E12", "Signature: ______________________________"]] as const;
  sheet.mergeCells("A10:C10"); sheet.mergeCells("A11:C11"); sheet.mergeCells("A12:C12"); sheet.mergeCells("E10:G10"); sheet.mergeCells("E11:G11"); sheet.mergeCells("E12:G12");
  details.forEach(([address, value]) => { const cell = sheet.getCell(address); cell.value = value; cell.font = blackFont; cell.alignment = left; });
  sheet.getRow(13).height = 8;

  const sections: Array<{ title: string; rows: Array<TransferCase | TransferAdminEntry> }> = [
    { title: "A. CRIMINAL CASES", rows: report.items.filter((item) => item.section === "Criminal").map((item) => item.snapshot) },
    { title: "B. CIVIL CASES", rows: report.items.filter((item) => item.section === "Civil").map((item) => item.snapshot) },
    ...(["ADMIN 1", "ADMIN 2", "ADMIN 3"] as const).map((section) => ({ title: section, rows: report.administrative_entries.filter((entry) => entry.section === section) })),
  ];
  let rowNumber = 15;
  sections.forEach((section) => {
    sheet.mergeCells(rowNumber, 1, rowNumber, 7);
    const sectionCell = sheet.getCell(rowNumber, 1); sectionCell.value = section.title; sectionCell.font = { ...blackFont, bold: true, size: 10 }; sectionCell.alignment = left; sectionCell.border = allBorders; sheet.getRow(rowNumber).height = 18; rowNumber += 1;
    const headerRow = sheet.getRow(rowNumber); headings.forEach((heading, index) => { const cell = headerRow.getCell(index + 1); cell.value = heading; cell.font = { ...blackFont, bold: true, size: 8 }; cell.alignment = center; cell.border = allBorders; }); headerRow.height = 28; rowNumber += 1;
    if (!section.rows.length) {
      const emptyRow = sheet.getRow(rowNumber); emptyRow.getCell(3).value = "No entries"; for (let index = 1; index <= 7; index += 1) { const cell = emptyRow.getCell(index); cell.font = blackFont; cell.alignment = center; cell.border = allBorders; } emptyRow.height = 20; rowNumber += 1;
    } else {
      section.rows.forEach((row, index) => {
        const values = row as unknown as Record<string, unknown>;
        const data = [index + 1, values.control_no ?? "", values.title_of_case ?? "", values.case_no ?? values.docket_no ?? "", formatTransferDate(String(values.date_received || "")), values.status_of_case ?? "", values.cause_of_action ?? ""];
        const dataRow = sheet.getRow(rowNumber); data.forEach((value, columnIndex) => { const cell = dataRow.getCell(columnIndex + 1); cell.value = value as string | number; cell.font = blackFont; cell.alignment = columnIndex === 0 ? center : left; cell.border = allBorders; }); dataRow.height = Math.max(18, Math.min(72, 18 * Math.max(1, String(data[2]).length / 42, String(data[6]).length / 44))); rowNumber += 1;
      });
    }
    rowNumber += 1;
  });
  sheet.pageSetup.printArea = `A1:G${rowNumber - 1}`;
  sheet.pageSetup.printTitlesRow = "15:16";
  sheet.headerFooter.oddFooter = "&CPage &P of &N";
  const [logo1Buffer, logo2Buffer] = await Promise.all([loadImage(logo1Url), loadImage(logo2Url)]);
  const logo1Id = workbook.addImage({ buffer: logo1Buffer, extension: "png" });
  const logo2Id = workbook.addImage({ buffer: logo2Buffer, extension: "png" });
  sheet.addImage(logo1Id, { tl: { col: 0.15, row: 0.15 }, ext: { width: 62, height: 62 } });
  sheet.addImage(logo2Id, { tl: { col: 5.95, row: 0.15 }, ext: { width: 82, height: 62 } });
  return workbook.xlsx.writeBuffer();
}
