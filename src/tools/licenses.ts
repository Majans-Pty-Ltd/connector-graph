import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { GraphClient } from "../api/client.js";
import type { GraphSubscribedSku, GraphLicenseDetail, ODataResponse } from "../api/types.js";

export function registerLicenseTools(server: McpServer, client: GraphClient): void {
  server.tool(
    "graph_list_subscribed_skus",
    "List all subscribed license SKUs with consumed/prepaid counts. Shows license inventory for the tenant.",
    {},
    async () => {
      try {
        const result = await client.get<ODataResponse<GraphSubscribedSku>>("subscribedSkus");
        const skus = result.value.map((sku) => ({
          skuPartNumber: sku.skuPartNumber,
          skuId: sku.skuId,
          capabilityStatus: sku.capabilityStatus,
          consumedUnits: sku.consumedUnits,
          enabled: sku.prepaidUnits.enabled,
          suspended: sku.prepaidUnits.suspended,
          available: sku.prepaidUnits.enabled - sku.consumedUnits,
        }));
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({ total: skus.length, skus }, null, 2),
            },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );

  server.tool(
    "graph_list_user_licenses",
    "List license details for a specific user.",
    {
      user_id: z.string().describe("User ID (GUID) or UPN"),
    },
    async ({ user_id }) => {
      try {
        const result = await client.get<ODataResponse<GraphLicenseDetail>>(
          `users/${encodeURIComponent(user_id)}/licenseDetails`
        );
        const licenses = result.value.map((lic) => ({
          skuPartNumber: lic.skuPartNumber,
          skuId: lic.skuId,
          servicePlans: lic.servicePlans
            .filter((sp) => sp.provisioningStatus === "Success")
            .map((sp) => sp.servicePlanName),
        }));
        return {
          content: [
            {
              type: "text" as const,
              text: JSON.stringify({ user: user_id, total: licenses.length, licenses }, null, 2),
            },
          ],
        };
      } catch (err) {
        return {
          content: [{ type: "text" as const, text: `Error: ${(err as Error).message}` }],
          isError: true,
        };
      }
    }
  );
}
