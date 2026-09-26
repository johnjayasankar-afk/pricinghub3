/* PricingHub, in the Labs material.
   Glass on the bar, the result tiles, the download rows and the ghost button.
   The dashboard screenshots are left exactly as they are: they are the product. */
import { initLabsUI } from './labs-ui.js';

initLabsUI({
  glass: [
    { sel: 'header.top', spec: 1, lens: [13, 52, 9, 1.95], vars: { '--gl-tint': '.52' } },
    { sel: '.tile', lens: [12, 34, 8, 1.7], vars: { '--gl-tint': '.66' } },
    { sel: '.doc', spec: 1, lens: [12, 34, 8, 1.7], vars: { '--gl-tint': '.62' } },
    { sel: '.btn-ghost', spec: 1, vars: { '--gl-tint': '.14' } }
  ],
  headings: 'h1, h2'
});
