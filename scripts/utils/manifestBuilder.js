'use strict';

const fs = require('fs-extra');

function unique(items) {
  return Array.from(new Set(items.filter(Boolean)));
}

function createFirefoxManifest({
  chromiumManifestPath,
  firefoxManifestTemplatePath,
  outputManifestPath,
}) {
  const chromiumManifest = fs.readJsonSync(chromiumManifestPath);
  const firefoxTemplate = fs.readJsonSync(firefoxManifestTemplatePath);

  const action = chromiumManifest.action || {};

  const firefoxManifest = {
    ...firefoxTemplate,
    // Keep shared extension identity fields in one place (public/manifest.json).
    version: chromiumManifest.version,
    name: chromiumManifest.name,
    short_name: chromiumManifest.short_name,
    description: chromiumManifest.description,
    icons: chromiumManifest.icons,
    content_scripts: chromiumManifest.content_scripts,
    devtools_page: chromiumManifest.devtools_page,
    browser_action: {
      ...(firefoxTemplate.browser_action || {}),
      default_title: action.default_title,
      default_popup: action.default_popup,
    },
  };

  const firefoxTemplatePermissions = Array.isArray(firefoxTemplate.permissions)
    ? firefoxTemplate.permissions
    : [];
  const chromiumPermissions = Array.isArray(chromiumManifest.permissions)
    ? chromiumManifest.permissions
    : [];
  const chromiumHostPermissions = Array.isArray(chromiumManifest.host_permissions)
    ? chromiumManifest.host_permissions
    : [];

  // Firefox MV2 stores host origins in permissions, so merge them from Chromium host_permissions.
  firefoxManifest.permissions = unique([
    ...firefoxTemplatePermissions,
    ...chromiumPermissions,
    ...chromiumHostPermissions,
  ]);

  fs.writeJsonSync(outputManifestPath, firefoxManifest, { spaces: 2 });
  return firefoxManifest;
}

// Append a suffix (e.g. " (DEV)") to the extension's display names so a DEV build
// is distinguishable from production and can be installed alongside it. Mutates
// the manifest file in place. Safe to call on the Chromium MV3 manifest before
// the Firefox manifest is derived from it, so both builds inherit the suffix.
function applyDevManifestLabel(manifestPath, suffix = ' (DEV)') {
  const manifest = fs.readJsonSync(manifestPath);

  if (typeof manifest.name === 'string') {
    manifest.name += suffix;
  }
  if (typeof manifest.short_name === 'string') {
    manifest.short_name += suffix;
  }
  if (manifest.action && typeof manifest.action.default_title === 'string') {
    manifest.action.default_title += suffix;
  }
  // Firefox MV2 keeps the action under browser_action.
  if (manifest.browser_action && typeof manifest.browser_action.default_title === 'string') {
    manifest.browser_action.default_title += suffix;
  }

  fs.writeJsonSync(manifestPath, manifest, { spaces: 2 });
  return manifest;
}

module.exports = {
  createFirefoxManifest,
  applyDevManifestLabel,
};