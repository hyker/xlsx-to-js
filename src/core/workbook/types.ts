import { WorkSheet } from "../worksheet/types";

/** 
 * A collection of worksheets. 
 */
export interface Workbook {
    date1904?: boolean;
    /** Define a sheet array and its contents. */
    workSheets: WorkSheet[];
}
