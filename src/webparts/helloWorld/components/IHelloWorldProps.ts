import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHelloWorldProps {
  description: string;
  slider: number;
  context: WebPartContext;
}
