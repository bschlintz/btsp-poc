import { ICommonResponse } from "../models/ICommonResponse";

export const JsonResponse = (status: number, body: ICommonResponse): any => {
  return {
    status,
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  }
}