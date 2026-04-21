import { parseIncomingEnquiry, type ParsedEnquiry } from "./enquiry-parser.js";
import {
  buildTracks,
  clientInputs,
  productReferences,
  productSheetReferences,
  quotationTemplateReferences
} from "./phase1-data.js";

const sampleMessages = [
  `8237116493
Paras from Nagpur, Maharashtra
send videos and detail details of mulberry silk reeling machine`,
  `Akash
Karnataka
Mulberry reeling machine`,
  `+91 69006 61188: Mangesh Kalwar
Lala-4, public School Road, 788163, Assam
Buniyaad Reeling Machine video to share`,
  `I want to know details of mulberry silk reeling machine
Shravani yamsani
Manasi silks
Hyderabad, 9849439912`
];

const parsedSamples = sampleMessages.map((message) => parseIncomingEnquiry(message));

function buildPipelinePreview(samples: ParsedEnquiry[]) {
  return [
    { stage: "New", count: samples.length },
    {
      stage: "Needs Review",
      count: samples.filter((sample) => sample.confidence === "low").length
    },
    {
      stage: "Quotation Draft",
      count: samples.filter((sample) => sample.confidence !== "low").length
    },
    { stage: "Quotation Sent", count: 0 },
    { stage: "Converted", count: 0 }
  ];
}

export function getProjectSummary() {
  return {
    metrics: [
      { label: "Sample enquiries mapped", value: sampleMessages.length },
      { label: "Reference products", value: productReferences.length },
      {
        label: "Quotation templates",
        value: quotationTemplateReferences.length
      },
      { label: "Product sheets", value: productSheetReferences.length },
      {
        label: "Inputs received",
        value: clientInputs.filter((item) => item.status === "received").length
      }
    ],
    pipeline: buildPipelinePreview(parsedSamples),
    buildTracks
  };
}

export function getProjectSnapshot() {
  return {
    recommendation: {
      hosting: "DigitalOcean VPS",
      automation: "Self-hosted n8n",
      crm: "Zoho Bigin",
      workspace: "Airtable",
      storage: "Google Drive",
      whatsapp: "Meta WhatsApp Cloud API"
    },
    quotationTemplateReferences,
    productSheetReferences,
    productReferences,
    clientInputs,
    buildTracks,
    parsedSamples
  };
}
