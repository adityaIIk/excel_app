export interface ClaimData {
  claimNumber?: string;
  policyNumber?: string;
  insuredName?: string;
  insuredPhone?: string;
  insuredEmail?: string;
  vehicleType?: string;
  financerName?: string;
  insurancePeriodFrom?: string;
  insurancePeriodTo?: string;
  idv?: string;

  regNumber?: string;
  regDate?: string;
  vehicleSaleDate?: string;
  chassisNumber?: string;
  engineNumber?: string;
  makeModel?: string;
  colour?: string;
  cubicCapacity?: string;
  seatingCapacity?: string;

  accidentDate?: string;
  accidentPlace?: string;
  dateOfIntimation?: string;
  estimatedLoss?: string;
  causeOfLoss?: string;

  driverName?: string;
  dlNumber?: string;
  dlValidUpto?: string;
  dlIssuedOn?: string;
  dlIssuingAuthority?: string;
  licenseType?: string;
}

export interface PartEntry {
  name: string;
  material: string;
  gst: string;
  estimatedCost: string;
  billedAmount: string;
}

export interface ExtractionResult {
  claimData: ClaimData;
  parts: PartEntry[];
}
