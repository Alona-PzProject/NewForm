import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INewFormProps {
  description: string;
  context: WebPartContext;
  webURL: string;
}
