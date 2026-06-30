import {
  generateLocalTerraformSnippet,
  normalizeTerraformUrl,
  stripReadOnlyFields,
  formatHclValue,
  formatHclString,
  buildResourceLabel,
  renderTerraformBlock,
  prependGetWarning,
  uniqueLabel,
  sanitizeLabel,
  shouldStripOdataKey,
  hashUrl,
  leafCollectionName,
  tryParseJson,
  parseODataContext,
  generateLocalTerraformBatchSnippet,
} from "./client.js";

const GUID = "12345678-1234-1234-1234-123456789012";

describe("tryParseJson", () => {
  it("parses a JSON object", () => {
    expect(tryParseJson('{"a":1}')).toEqual({ a: 1 });
  });

  it("parses a JSON array", () => {
    expect(tryParseJson("[1,2,3]")).toEqual([1, 2, 3]);
  });

  it("trims surrounding whitespace before parsing", () => {
    expect(tryParseJson('   {"a":1}   ')).toEqual({ a: 1 });
  });

  it("returns null for whitespace-only input", () => {
    expect(tryParseJson("   ")).toBeNull();
  });

  it("returns null for an empty string", () => {
    expect(tryParseJson("")).toBeNull();
  });

  it("returns null for null / non-string input", () => {
    expect(tryParseJson(null)).toBeNull();
    expect(tryParseJson(undefined)).toBeNull();
    expect(tryParseJson(123)).toBeNull();
    expect(tryParseJson({})).toBeNull();
  });

  it("returns null (does not throw) for malformed JSON", () => {
    expect(tryParseJson("{not json}")).toBeNull();
  });
});

describe("hashUrl", () => {
  it("is deterministic for the same input", () => {
    expect(hashUrl("https://graph.microsoft.com/v1.0/groups")).toBe(
      hashUrl("https://graph.microsoft.com/v1.0/groups")
    );
  });

  it("produces a short alphanumeric string of at most 6 chars", () => {
    expect(hashUrl("some-seed-value")).toMatch(/^[a-z0-9]{1,6}$/);
  });

  it("produces different output for different input", () => {
    expect(hashUrl("seed-a")).not.toBe(hashUrl("seed-b"));
  });

  it("handles the empty string", () => {
    expect(hashUrl("")).toBe("0");
  });
});

describe("sanitizeLabel", () => {
  it("converts camelCase to snake_case", () => {
    expect(sanitizeLabel("displayName")).toBe("display_name");
  });

  it("replaces spaces and symbols with underscores", () => {
    expect(sanitizeLabel("Test Group!")).toBe("test_group");
  });

  it("strips leading and trailing underscores", () => {
    expect(sanitizeLabel("  hello  ")).toBe("hello");
    expect(sanitizeLabel("@@name@@")).toBe("name");
  });

  it("lowercases and keeps digits", () => {
    expect(sanitizeLabel("Group123")).toBe("group123");
  });

  it("collapses runs of invalid characters into a single underscore", () => {
    expect(sanitizeLabel("a---b...c")).toBe("a_b_c");
  });

  it("coerces non-string input via String()", () => {
    expect(sanitizeLabel(42)).toBe("42");
  });

  it("returns an empty string when nothing valid remains", () => {
    expect(sanitizeLabel("***")).toBe("");
  });
});

describe("shouldStripOdataKey", () => {
  it("matches an exact strip key", () => {
    expect(shouldStripOdataKey("@odata.context")).toBe(true);
    expect(shouldStripOdataKey("@odata.etag")).toBe(true);
  });

  it("matches a property-scoped suffix (foo@odata.etag)", () => {
    expect(shouldStripOdataKey("photo@odata.mediaEtag")).toBe(true);
  });

  it("keeps @odata.type discriminators", () => {
    expect(shouldStripOdataKey("@odata.type")).toBe(false);
  });

  it("keeps *@odata.bind references", () => {
    expect(shouldStripOdataKey("members@odata.bind")).toBe(false);
    expect(shouldStripOdataKey("@odata.bind")).toBe(false);
  });

  it("returns false for ordinary keys", () => {
    expect(shouldStripOdataKey("displayName")).toBe(false);
  });
});

describe("stripReadOnlyFields", () => {
  it("removes top-level read-only keys", () => {
    const input = {
      id: GUID,
      createdDateTime: "2020-01-01",
      deletedDateTime: "2020-01-02",
      renewedDateTime: "2020-01-03",
      modifiedDateTime: "2020-01-04",
      displayName: "keep",
    };
    expect(stripReadOnlyFields(input)).toEqual({ displayName: "keep" });
  });

  it("keeps read-only keys when nested (depth > 0)", () => {
    const input = { owner: { id: "keep-me", displayName: "o" } };
    expect(stripReadOnlyFields(input)).toEqual({
      owner: { id: "keep-me", displayName: "o" },
    });
  });

  it("strips @odata.* annotation keys at any depth", () => {
    const input = {
      "@odata.context": "ctx",
      nested: { "@odata.etag": "e", name: "n" },
    };
    expect(stripReadOnlyFields(input)).toEqual({ nested: { name: "n" } });
  });

  it("preserves @odata.type and @odata.bind at any depth", () => {
    const input = {
      "@odata.type": "#microsoft.graph.user",
      nested: { "members@odata.bind": ["url"] },
    };
    expect(stripReadOnlyFields(input)).toEqual({
      "@odata.type": "#microsoft.graph.user",
      nested: { "members@odata.bind": ["url"] },
    });
  });

  it("recurses through arrays of objects", () => {
    const input = { items: [{ "@odata.etag": "x", v: 1 }, { v: 2 }] };
    expect(stripReadOnlyFields(input)).toEqual({ items: [{ v: 1 }, { v: 2 }] });
  });

  it("passes primitives and null through unchanged", () => {
    expect(stripReadOnlyFields("str")).toBe("str");
    expect(stripReadOnlyFields(7)).toBe(7);
    expect(stripReadOnlyFields(null)).toBeNull();
  });
});

describe("formatHclString", () => {
  it("wraps a plain string in quotes", () => {
    expect(formatHclString("hello")).toBe('"hello"');
  });

  it("escapes backslashes, quotes, newlines, carriage returns and tabs", () => {
    expect(formatHclString('a\\b"c\nd\re\tf')).toBe('"a\\\\b\\"c\\nd\\re\\tf"');
  });

  it("coerces non-string input", () => {
    expect(formatHclString(5)).toBe('"5"');
  });
});

describe("formatHclValue", () => {
  it("renders null and undefined as null", () => {
    expect(formatHclValue(null, "")).toBe("null");
    expect(formatHclValue(undefined, "")).toBe("null");
  });

  it("renders booleans", () => {
    expect(formatHclValue(true, "")).toBe("true");
    expect(formatHclValue(false, "")).toBe("false");
  });

  it("renders numbers as-is", () => {
    expect(formatHclValue(42, "")).toBe("42");
  });

  it("delegates strings to HCL escaping", () => {
    expect(formatHclValue('say "hi"', "")).toBe('"say \\"hi\\""');
  });

  it("renders an empty array and empty object compactly", () => {
    expect(formatHclValue([], "")).toBe("[]");
    expect(formatHclValue({}, "")).toBe("{}");
  });

  it("renders a non-empty array with indentation", () => {
    expect(formatHclValue(["a", "b"], "")).toBe('[\n  "a",\n  "b",\n]');
  });

  it("aligns object keys by width", () => {
    expect(formatHclValue({ a: 1, bbb: 2 }, "")).toBe(
      "{\n  a   = 1\n  bbb = 2\n}"
    );
  });

  it("quotes non-identifier keys and leaves identifier keys bare", () => {
    const out = formatHclValue({ "@odata.type": "#x", name: "n" }, "");
    expect(out).toContain('"@odata.type"');
    expect(out).toContain('name');
    expect(out).not.toContain('"name"');
  });

  it("renders nested structures", () => {
    const out = formatHclValue({ outer: { inner: [1] } }, "");
    expect(out).toBe("{\n  outer = {\n    inner = [\n      1,\n    ]\n  }\n}");
  });
});

describe("normalizeTerraformUrl", () => {
  it("strips the v1.0 prefix and reports no api version", () => {
    expect(
      normalizeTerraformUrl("https://graph.microsoft.com/v1.0/groups")
    ).toEqual({ url: "groups", apiVersion: null });
  });

  it("detects the beta prefix as an api version", () => {
    expect(
      normalizeTerraformUrl("https://graph.microsoft.com/beta/groups")
    ).toEqual({ url: "groups", apiVersion: "beta" });
  });

  it("matches the version prefix case-insensitively", () => {
    expect(
      normalizeTerraformUrl("https://graph.microsoft.com/V1.0/groups")
    ).toEqual({ url: "groups", apiVersion: null });
  });

  it("drops a trailing GUID when there is more than one segment", () => {
    expect(
      normalizeTerraformUrl(`https://graph.microsoft.com/v1.0/groups/${GUID}`)
    ).toEqual({ url: "groups", apiVersion: null });
  });

  it("keeps a lone GUID segment (only trims when >1 segment)", () => {
    expect(
      normalizeTerraformUrl(`https://graph.microsoft.com/v1.0/${GUID}`)
    ).toEqual({ url: GUID, apiVersion: null });
  });

  it("removes the query string", () => {
    expect(
      normalizeTerraformUrl(
        "https://graph.microsoft.com/v1.0/groups?$select=id,displayName"
      )
    ).toEqual({ url: "groups", apiVersion: null });
  });

  it("handles a relative path with no recognised domain", () => {
    expect(
      normalizeTerraformUrl(`/v1.0/users/${GUID}`)
    ).toEqual({ url: "users", apiVersion: null });
  });

  it("preserves nested resource paths", () => {
    expect(
      normalizeTerraformUrl(
        `https://graph.microsoft.com/v1.0/groups/${GUID}/members`
      )
    ).toEqual({ url: `groups/${GUID}/members`, apiVersion: null });
  });
});

describe("leafCollectionName", () => {
  it("returns the single segment", () => {
    expect(leafCollectionName("groups")).toBe("groups");
  });

  it("returns the last non-GUID segment", () => {
    expect(leafCollectionName(`groups/${GUID}/members`)).toBe("members");
  });

  it("skips a trailing GUID", () => {
    expect(leafCollectionName(`groups/${GUID}`)).toBe("groups");
  });

  it("ignores $-prefixed segments", () => {
    expect(leafCollectionName("groups/$count")).toBe("groups");
  });

  it("returns empty string when there is no collection", () => {
    expect(leafCollectionName("")).toBe("");
  });
});

describe("buildResourceLabel", () => {
  const singulars = {
    groups: "group",
    users: "user",
    applications: "application",
    servicePrincipals: "service_principal",
    oauth2PermissionGrants: "oauth2_permission_grant",
    appRoleAssignedTo: "app_role_assignment",
    federatedIdentityCredentials: "federated_identity_credential",
    devices: "device",
    contacts: "contact",
  };

  Object.entries(singulars).forEach(([collection, singular]) => {
    it(`maps the "${collection}" collection to the "${singular}" prefix`, () => {
      expect(buildResourceLabel(collection, { displayName: "X" }, "seed")).toBe(
        `${singular}_x`
      );
    });
  });

  it("falls back to sanitizeLabel for an unknown collection", () => {
    expect(
      buildResourceLabel("widgets", { displayName: "My Thing" }, "seed")
    ).toBe("widgets_my_thing"); // sanitizeLabel("widgets") === "widgets"
  });

  it("uses the 'resource' prefix when no collection name is present", () => {
    expect(buildResourceLabel("", {}, "seed")).toBe(`resource_${hashUrl("seed")}`);
  });

  it("prefers displayName over the other name candidates", () => {
    expect(
      buildResourceLabel(
        "groups",
        { displayName: "Primary", mailNickname: "nick", name: "n" },
        "seed"
      )
    ).toBe("group_primary");
  });

  it("falls through the candidate priority: mailNickname, then userPrincipalName, then name", () => {
    expect(buildResourceLabel("groups", { mailNickname: "nick" }, "s")).toBe(
      "group_nick"
    );
    expect(
      buildResourceLabel("users", { userPrincipalName: "upn@contoso" }, "s")
    ).toBe("user_upn_contoso");
    expect(buildResourceLabel("contacts", { name: "Bob" }, "s")).toBe(
      "contact_bob"
    );
  });

  it("skips empty / whitespace-only candidates", () => {
    expect(
      buildResourceLabel("groups", { displayName: "   ", mailNickname: "nick" }, "s")
    ).toBe("group_nick");
  });

  it("uses a deterministic hash fallback when no name candidate is found", () => {
    const seed = "https://graph.microsoft.com/v1.0/groups#0";
    expect(buildResourceLabel("groups", { foo: "bar" }, seed)).toBe(
      `group_${hashUrl(seed)}`
    );
    expect(buildResourceLabel("groups", { foo: "bar" }, seed)).toBe(
      buildResourceLabel("groups", { foo: "bar" }, seed)
    );
  });
});

describe("renderTerraformBlock", () => {
  it("renders url and body and omits api_version when not set", () => {
    const block = renderTerraformBlock("group_x", "groups", null, {
      displayName: "X",
    });
    expect(block).toBe(
      'resource "msgraph_resource" "group_x" {\n' +
        '  url = "groups"\n' +
        "  body = {\n" +
        '    displayName = "X"\n' +
        "  }\n" +
        "}"
    );
  });

  it("includes an api_version line when set", () => {
    const block = renderTerraformBlock("group_x", "groups", "beta", {
      displayName: "X",
    });
    expect(block).toContain('  api_version = "beta"');
  });
});

describe("prependGetWarning", () => {
  it("prepends the three warning comment lines above the block", () => {
    const result = prependGetWarning("BLOCK");
    const lines = result.split("\n");
    expect(lines[0]).toBe(
      "# WARNING: Generated from a GET response - review before apply."
    );
    expect(lines[1]).toMatch(/^# Read-only/);
    expect(lines[2]).toMatch(/^# verify/);
    expect(lines[3]).toBe("BLOCK");
  });
});

describe("uniqueLabel", () => {
  it("returns the base label on first use and records it", () => {
    const used = new Set();
    expect(uniqueLabel("group_x", used)).toBe("group_x");
    expect(used.has("group_x")).toBe(true);
  });

  it("suffixes _2, _3 on subsequent collisions", () => {
    const used = new Set();
    expect(uniqueLabel("group_x", used)).toBe("group_x");
    expect(uniqueLabel("group_x", used)).toBe("group_x_2");
    expect(uniqueLabel("group_x", used)).toBe("group_x_3");
  });
});

describe("generateLocalTerraformSnippet", () => {
  const URL = "https://graph.microsoft.com/v1.0/groups";

  it("returns null for DELETE and OPTIONS", () => {
    expect(generateLocalTerraformSnippet("DELETE", URL, "{}", "")).toBeNull();
    expect(generateLocalTerraformSnippet("OPTIONS", URL, "{}", "")).toBeNull();
  });

  it("returns null for GET without the experimental flag", () => {
    expect(
      generateLocalTerraformSnippet("GET", URL, "", '{"displayName":"X"}')
    ).toBeNull();
  });

  it("generates a block from a GET response when the experimental flag is set, with a warning", () => {
    const out = generateLocalTerraformSnippet(
      "GET",
      URL,
      "",
      `{"id":"${GUID}","displayName":"X"}`,
      true
    );
    expect(out).toContain("# WARNING: Generated from a GET response");
    expect(out).toContain('resource "msgraph_resource" "group_x"');
    expect(out).not.toContain(GUID); // read-only id stripped
  });

  it("produces an exact HCL block for a POST create", () => {
    const out = generateLocalTerraformSnippet(
      "POST",
      URL,
      '{"displayName":"Marketing"}',
      ""
    );
    expect(out).toBe(
      'resource "msgraph_resource" "group_marketing" {\n' +
        '  url = "groups"\n' +
        "  body = {\n" +
        '    displayName = "Marketing"\n' +
        "  }\n" +
        "}"
    );
  });

  it("does not add a warning for non-GET methods", () => {
    const out = generateLocalTerraformSnippet(
      "PATCH",
      URL,
      '{"displayName":"X"}',
      ""
    );
    expect(out).not.toContain("WARNING");
  });

  it("falls back to responseBody for non-GET when requestBody is empty", () => {
    const out = generateLocalTerraformSnippet("POST", URL, "", '{"displayName":"X"}');
    expect(out).toContain('resource "msgraph_resource" "group_x"');
  });

  it("falls back to requestBody for GET when responseBody is empty", () => {
    const out = generateLocalTerraformSnippet(
      "GET",
      URL,
      '{"displayName":"X"}',
      "",
      true
    );
    expect(out).toContain('resource "msgraph_resource" "group_x"');
  });

  it("returns null for invalid / empty JSON", () => {
    expect(generateLocalTerraformSnippet("POST", URL, "not json", "")).toBeNull();
    expect(generateLocalTerraformSnippet("POST", URL, "", "")).toBeNull();
  });

  it("returns null for primitive (non-object) JSON", () => {
    expect(generateLocalTerraformSnippet("POST", URL, "123", "")).toBeNull();
    expect(generateLocalTerraformSnippet("POST", URL, '"a string"', "")).toBeNull();
  });

  it("emits one block per item for a collection response", () => {
    const out = generateLocalTerraformSnippet(
      "GET",
      URL,
      "",
      '{"value":[{"displayName":"Alpha"},{"displayName":"Beta"}]}',
      true
    );
    const blocks = out.split("\n\n");
    expect(blocks).toHaveLength(2);
    expect(out).toContain('"group_alpha"');
    expect(out).toContain('"group_beta"');
  });

  it("strips read-only and @odata fields from output", () => {
    const out = generateLocalTerraformSnippet(
      "POST",
      URL,
      `{"@odata.context":"ctx","id":"${GUID}","displayName":"X"}`,
      ""
    );
    expect(out).not.toContain("@odata.context");
    expect(out).not.toContain(GUID);
    expect(out).toContain('displayName = "X"');
  });

  it("returns null when every item is empty after stripping", () => {
    expect(generateLocalTerraformSnippet("POST", URL, `{"id":"${GUID}"}`, "")).toBeNull();
  });

  it("de-duplicates labels across collection items", () => {
    const out = generateLocalTerraformSnippet(
      "GET",
      URL,
      "",
      '{"value":[{"displayName":"Dup"},{"displayName":"Dup"}]}',
      true
    );
    expect(out).toContain('"group_dup"');
    expect(out).toContain('"group_dup_2"');
  });

  it("emits api_version for a beta URL and omits it for v1.0", () => {
    const beta = generateLocalTerraformSnippet(
      "POST",
      "https://graph.microsoft.com/beta/groups",
      '{"displayName":"X"}',
      ""
    );
    expect(beta).toContain('api_version = "beta"');

    const v1 = generateLocalTerraformSnippet("POST", URL, '{"displayName":"X"}', "");
    expect(v1).not.toContain("api_version");
  });
});

describe("parseODataContext", () => {
  it("parses a collection context into path + beta api version", () => {
    expect(
      parseODataContext(
        "https://graph.microsoft.com/beta/$metadata#identity/identityProviders"
      )
    ).toEqual({ url: "identity/identityProviders", apiVersion: "beta" });
  });

  it("reports null api version for a v1.0 context", () => {
    expect(
      parseODataContext("https://graph.microsoft.com/v1.0/$metadata#groups")
    ).toEqual({ url: "groups", apiVersion: null });
  });

  it("strips a trailing /$entity single-entity marker", () => {
    expect(
      parseODataContext("https://graph.microsoft.com/v1.0/$metadata#groups/$entity")
    ).toEqual({ url: "groups", apiVersion: null });
  });

  it("strips projection and key-selector parens", () => {
    expect(
      parseODataContext(
        "https://graph.microsoft.com/v1.0/$metadata#users(id,displayName)"
      )
    ).toEqual({ url: "users", apiVersion: null });
    expect(
      parseODataContext("https://graph.microsoft.com/v1.0/$metadata#groups('abc')")
    ).toEqual({ url: "groups", apiVersion: null });
  });

  it("returns null when the $metadata marker is missing", () => {
    expect(parseODataContext("https://graph.microsoft.com/v1.0/groups")).toBeNull();
  });

  it("returns null for non-string input", () => {
    expect(parseODataContext(null)).toBeNull();
    expect(parseODataContext(undefined)).toBeNull();
    expect(parseODataContext(42)).toBeNull();
  });

  it("returns null when the fragment is empty", () => {
    expect(parseODataContext("https://graph.microsoft.com/v1.0/$metadata#")).toBeNull();
  });
});

// The real $batch response payload provided for this feature.
const IDENTITY_PROVIDERS_BATCH = JSON.stringify({
  responses: [
    {
      id: "38e548c6-045f-4a4a-9b18-b8d8644cf94e",
      status: 200,
      headers: { "Content-Type": "application/json" },
      body: {
        "@odata.context":
          "https://graph.microsoft.com/beta/$metadata#identity/identityProviders",
        value: [
          {
            "@odata.type": "#microsoft.graph.socialIdentityProvider",
            id: "5a415368-09cf-42df-a00c-ec23d9f4ebce",
            displayName: "Facebook",
            supportedTenantTypes: "externalId",
            identityProviderType: "Facebook",
            clientId: "1515151515151515",
            clientSecret: "******",
          },
          {
            "@odata.type": "#microsoft.graph.oidcIdentityProvider",
            id: "6d5887a6-131d-44a6-95ed-ff8c2cd57f42",
            displayName: "WorkforceEntraId",
            supportedTenantTypes: "externalId",
            clientId: "a9a9a9a9-a9a9-a9a9-a9a9-a9a9a9a9a9a9",
            issuer:
              "https://login.microsoftonline.com/c5c5c5c5-c5c5-c5c5-c5c5-c5c5c5c5c5c5/v2.0",
            responseType: "code",
            scope: "openid profile",
            clientAuthentication: {
              "@odata.type": "#microsoft.graph.oidcClientSecretAuthentication",
              clientSecret: "******",
            },
          },
        ],
      },
    },
  ],
});

describe("generateLocalTerraformBatchSnippet", () => {
  it("generates one adopt block per resource in the provided identityProviders payload", () => {
    const out = generateLocalTerraformBatchSnippet("", IDENTITY_PROVIDERS_BATCH);
    const blocks = out.split("\n\n");
    expect(blocks).toHaveLength(2);
  });

  it("derives url and beta api_version from @odata.context", () => {
    const out = generateLocalTerraformBatchSnippet("", IDENTITY_PROVIDERS_BATCH);
    expect(out).toContain('url = "identity/identityProviders"');
    expect(out).toContain('api_version = "beta"');
  });

  it("preserves @odata.type discriminators and strips read-only ids", () => {
    const out = generateLocalTerraformBatchSnippet("", IDENTITY_PROVIDERS_BATCH);
    // (HCL aligns the `=`, so match the quoted key + value without assuming spacing)
    expect(out).toContain('"@odata.type"');
    expect(out).toContain('"#microsoft.graph.oidcIdentityProvider"');
    expect(out).toContain('"#microsoft.graph.socialIdentityProvider"');
    // nested discriminator preserved
    expect(out).toContain('"#microsoft.graph.oidcClientSecretAuthentication"');
    expect(out).not.toContain("5a415368-09cf-42df-a00c-ec23d9f4ebce"); // top-level id stripped
  });

  it("keeps non-read-only fields such as clientSecret", () => {
    const out = generateLocalTerraformBatchSnippet("", IDENTITY_PROVIDERS_BATCH);
    expect(out).toMatch(/clientSecret\s+=\s+"\*{6}"/);
  });

  it("prepends a GET/adopt warning to every block", () => {
    const out = generateLocalTerraformBatchSnippet("", IDENTITY_PROVIDERS_BATCH);
    const warnings = out.match(/# WARNING: Generated from a GET response/g);
    expect(warnings).toHaveLength(2);
  });

  it("derives labels from displayName", () => {
    const out = generateLocalTerraformBatchSnippet("", IDENTITY_PROVIDERS_BATCH);
    expect(out).toContain('"identity_providers_facebook"');
    expect(out).toContain('"identity_providers_workforce_entra_id"');
  });

  it("returns null for an invalid or non-batch response envelope", () => {
    expect(generateLocalTerraformBatchSnippet("", "not json")).toBeNull();
    expect(generateLocalTerraformBatchSnippet("", "{}")).toBeNull();
    expect(generateLocalTerraformBatchSnippet("", '{"value":[]}')).toBeNull();
  });

  it("skips non-2xx sub-responses", () => {
    const envelope = JSON.stringify({
      responses: [
        {
          id: "1",
          status: 404,
          body: {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups/$entity",
            displayName: "Missing",
          },
        },
        {
          id: "2",
          status: 200,
          body: {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups/$entity",
            displayName: "Present",
          },
        },
      ],
    });
    const out = generateLocalTerraformBatchSnippet("", envelope);
    expect(out).toContain('"group_present"');
    expect(out).not.toContain("Missing");
  });

  it("handles a single-entity body (no value array)", () => {
    const envelope = JSON.stringify({
      responses: [
        {
          id: "1",
          status: 200,
          body: {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups/$entity",
            id: GUID,
            displayName: "Solo",
          },
        },
      ],
    });
    const out = generateLocalTerraformBatchSnippet("", envelope);
    expect(out.split("\n\n")).toHaveLength(1);
    expect(out).toContain('"group_solo"');
    expect(out).toContain('url = "groups"');
    expect(out).not.toContain(GUID);
  });

  it("falls back to the paired request url when @odata.context is absent", () => {
    const requestBody = JSON.stringify({
      requests: [{ id: "42", method: "GET", url: "/groups" }],
    });
    const responseBody = JSON.stringify({
      responses: [{ id: "42", status: 200, body: { displayName: "FromReq" } }],
    });
    const out = generateLocalTerraformBatchSnippet(requestBody, responseBody);
    expect(out).toContain('url = "groups"');
    expect(out).toContain('"group_from_req"'); // sanitizeLabel("FromReq") -> from_req
  });

  it("skips responses with no resolvable url and no body fields", () => {
    const responseBody = JSON.stringify({
      responses: [
        { id: "1", status: 200, body: { displayName: "NoContext" } },
      ],
    });
    // no @odata.context and no matching request => skipped
    expect(generateLocalTerraformBatchSnippet("", responseBody)).toBeNull();
  });

  it("skips an item that is empty after stripping read-only fields", () => {
    const responseBody = JSON.stringify({
      responses: [
        {
          id: "1",
          status: 200,
          body: {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups/$entity",
            id: GUID,
          },
        },
      ],
    });
    expect(generateLocalTerraformBatchSnippet("", responseBody)).toBeNull();
  });

  it("de-duplicates labels across responses", () => {
    const responseBody = JSON.stringify({
      responses: [
        {
          id: "1",
          status: 200,
          body: {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups/$entity",
            displayName: "Dup",
          },
        },
        {
          id: "2",
          status: 200,
          body: {
            "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups/$entity",
            displayName: "Dup",
          },
        },
      ],
    });
    const out = generateLocalTerraformBatchSnippet("", responseBody);
    expect(out).toContain('"group_dup"');
    expect(out).toContain('"group_dup_2"');
  });
});

describe("generateLocalTerraformSnippet ($batch integration)", () => {
  const BATCH_URL = "https://graph.microsoft.com/beta/$batch";

  it("produces adopt blocks for a $batch when the experimental flag is set", () => {
    const out = generateLocalTerraformSnippet(
      "POST",
      BATCH_URL,
      "",
      IDENTITY_PROVIDERS_BATCH,
      true
    );
    expect(out.split("\n\n")).toHaveLength(2);
    expect(out).toContain('url = "identity/identityProviders"');
  });

  it("returns null for a $batch without the experimental flag", () => {
    expect(
      generateLocalTerraformSnippet("POST", BATCH_URL, "", IDENTITY_PROVIDERS_BATCH)
    ).toBeNull();
  });

  it("does not render the request envelope as a resource", () => {
    const requestBody = JSON.stringify({
      requests: [{ id: "1", method: "GET", url: "/identity/identityProviders" }],
    });
    const out = generateLocalTerraformSnippet(
      "POST",
      BATCH_URL,
      requestBody,
      IDENTITY_PROVIDERS_BATCH,
      true
    );
    expect(out).not.toContain('"requests"');
    expect(out).not.toContain("$batch");
  });
});
