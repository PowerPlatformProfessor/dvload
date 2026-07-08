// Tiny prompt helpers. We avoid pulling in a dependency for this — Node's
// readline + setRawMode is enough for the two interactions we need.

import * as readline from "node:readline/promises";
import { stdin, stdout } from "node:process";

/** Plain echoed prompt. Returns trimmed input. */
export async function prompt(question: string): Promise<string> {
  const rl = readline.createInterface({ input: stdin, output: stdout });
  try {
    const answer = await rl.question(question);
    return answer.trim();
  } finally {
    rl.close();
  }
}

/**
 * Prompt without echoing keystrokes. Used for client secrets.
 * Falls back to a regular prompt if stdin isn't a TTY (CI scenarios).
 */
export async function promptSecret(question: string): Promise<string> {
  if (!stdin.isTTY || !stdin.setRawMode) {
    return prompt(question);
  }

  return new Promise<string>((resolve, reject) => {
    stdout.write(question);
    stdin.setRawMode(true);
    stdin.resume();
    stdin.setEncoding("utf8");

    let buffer = "";
    const onData = (chunk: string) => {
      for (const ch of chunk) {
        const code = ch.charCodeAt(0);
        if (ch === "\n" || ch === "\r") {
          stdout.write("\n");
          cleanup();
          resolve(buffer);
          return;
        } else if (code === 3) {
          // Ctrl-C
          cleanup();
          reject(new Error("Cancelled"));
          return;
        } else if (code === 127 || code === 8) {
          // Backspace / DEL
          buffer = buffer.slice(0, -1);
        } else {
          buffer += ch;
        }
      }
    };

    function cleanup() {
      stdin.setRawMode(false);
      stdin.pause();
      stdin.removeListener("data", onData);
    }

    stdin.on("data", onData);
  });
}
