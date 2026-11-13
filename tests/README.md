# Review desk checks

Behavioural checks for the `review-desk` bundle. They run outside the app and do
not replace reading the panel in a browser.

```sh
cd tests
npm install          # react@18, react-dom@18, jsdom@24 (already present here)
node host-test.mjs   # route frames, prompts, web search, tool, error paths
node client-test.mjs # mounts the page in jsdom and drives the four actions
```

- `host-test.mjs` imports `../index.js` and installs it on a fake
  Cordis context whose `llm` streams canned chunks and whose `web` returns canned
  sources. It asserts the NDJSON frame order, evidence numbering and
  deduplication, payload validation, the agent tool result, the failure frames
  (model error, missing default model, malformed JSON), an explicit
  provider/model/reasoning route, refusal of an unknown route, both catalogue
  routes, and the custom-instruction action.
- `client-test.mjs` reads `../client.js`, captures the factory that
  the Conductor module loader receives, and mounts the registered page component
  in jsdom. It covers the empty state, the sample document, selection-driven
  enablement, streamed results, the reference list, the expansion replacement,
  mapped errors, the ⌘1–⌘4 shortcuts, a run still streaming, restoring a saved
  session, importing a stored-entry `.docx`, importing a `.pdf` text layer,
  rejecting a PDF without one, marking reviewed passages, locating a passage
  from its result card, and dropping a mark once the text no longer matches.

Selection is driven with `mouseup`, the event a completed drag produces; a
synthetic `select` event does not reach React's delegated handler under jsdom.
