import { createInterface } from "node:readline";

/** Asks questions one line at a time; closing releases the input stream. */
export type Prompter = {
  ask: (question: string) => Promise<string>;
  close: () => void;
};

/**
 * Creates a prompter over `input`. Answers are read from readline's async
 * iterator, which buffers lines that arrive before their question is asked, so
 * answers piped in all at once reach every question instead of only the first.
 * Input that ends before a question is answered is an error, not a hang.
 */
export const createPrompter = (
  input: NodeJS.ReadableStream,
  output: NodeJS.WritableStream,
): Prompter => {
  const rl = createInterface({ input, output, terminal: false });
  const lines = rl[Symbol.asyncIterator]();
  return {
    ask: async (question) => {
      output.write(question);
      const next = await lines.next();
      if (next.done) {
        throw new Error(`Input ended before an answer to: ${question.trim()}`);
      }
      return next.value;
    },
    close: () => rl.close(),
  };
};
