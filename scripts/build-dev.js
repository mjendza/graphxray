'use strict';

// Dedicated DEV build. Mark this as a DEV build before the real build reads the
// environment, then delegate to the standard production build. The flag is
// injected into the bundle via DefinePlugin (header "DEV Build" badge) and read
// again in build.js to append a "(DEV)" suffix to the extension manifest name.
process.env.REACT_APP_DEV_BUILD = 'true';

require('./build.js');
