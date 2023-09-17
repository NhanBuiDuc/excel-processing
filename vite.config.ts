import { sveltekit } from '@sveltejs/kit/vite';
import { defineConfig } from 'vite';

import { readFileSync } from 'fs'; // thêm vào

// Mới thêm vào
export default defineConfig({
  assetsInclude: ['**/*.numbers', '**/*.xlsx'],
  plugins: [sveltekit(), {
    name: "sheet-base64",
    transform(code, id) {
      if(!id.match(/\.(numbers|xlsx)$/)) return;
      var data = readFileSync(id, "base64");
      return `export default '${data}'`;
    }
  }]
});


// export default defineConfig({
// 	plugins: [
// 		sveltekit(),
// 	],

// });
