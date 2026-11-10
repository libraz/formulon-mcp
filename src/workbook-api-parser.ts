/** How `formulon_workbook_call` reaches a handle: its kind and the keys its accessors take, in order. */
export interface WorkbookApiTarget {
  kind: string;
  keys: string[];
}

export interface WorkbookApiMethod {
  name: string;
  target: WorkbookApiTarget;
  signature: string;
  description: string;
  access: "read" | "write";
}

export interface WorkbookApiDeclaration {
  name: string;
  kind: "interface" | "enum" | "type";
  description: string;
  declaration: string;
}

export type ParsedDeclaration = WorkbookApiDeclaration & {
  references: string[];
};

/** One method of an interface; overloads are merged, one signature per line. */
export type ParsedMethod = {
  name: string;
  signature: string;
  description: string;
  references: string[];
  /** The return type text of the first overload. */
  returns: string;
  /** The parameter count of the first overload. */
  arity: number;
};

/** The members of one exported interface, without those it inherits. */
export type ParsedInterface = {
  name: string;
  extends: string[];
  methods: Map<string, ParsedMethod>;
  /** Property name -> type text. */
  properties: Map<string, string>;
};

type Token = {
  value: string;
  start: number;
  end: number;
};

function scanTokens(source: string): Token[] {
  const tokens: Token[] = [];
  let index = 0;
  while (index < source.length) {
    const current = source[index];
    if (/\s/.test(current)) {
      index += 1;
      continue;
    }
    if (current === "/" && source[index + 1] === "/") {
      index += 2;
      while (index < source.length && source[index] !== "\n") {
        index += 1;
      }
      continue;
    }
    if (current === "/" && source[index + 1] === "*") {
      const end = source.indexOf("*/", index + 2);
      index = end === -1 ? source.length : end + 2;
      continue;
    }
    if (current === "'" || current === '"' || current === "`") {
      const quote = current;
      index += 1;
      while (index < source.length) {
        if (source[index] === "\\") {
          index += 2;
          continue;
        }
        if (source[index] === quote) {
          index += 1;
          break;
        }
        index += 1;
      }
      continue;
    }
    if (/[A-Za-z_$]/.test(current)) {
      const start = index;
      index += 1;
      while (index < source.length && /[A-Za-z0-9_$]/.test(source[index])) {
        index += 1;
      }
      tokens.push({ value: source.slice(start, index), start, end: index });
      continue;
    }
    tokens.push({ value: current, start: index, end: index + 1 });
    index += 1;
  }
  return tokens;
}

function matchingDelimiter(tokens: Token[], openIndex: number): number {
  const opening = tokens[openIndex]?.value;
  const closing = opening === "{" ? "}" : opening === "(" ? ")" : opening === "[" ? "]" : "";
  if (!closing) {
    throw new Error(`internal parser error: ${opening} is not an opening delimiter`);
  }
  let depth = 0;
  for (let index = openIndex; index < tokens.length; index += 1) {
    if (tokens[index].value === opening) {
      depth += 1;
    } else if (tokens[index].value === closing) {
      depth -= 1;
      if (depth === 0) {
        return index;
      }
    }
  }
  throw new Error(`unbalanced ${opening} delimiter in Formulon declaration file`);
}

function normalizeDoc(raw: string): string {
  const body = raw.replace(/^\/\*\*?/, "").replace(/\*\/$/, "");
  return body
    .split("\n")
    .map((line) => line.replace(/^\s*\* ?/, "").trim())
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function leadingDoc(source: string, start: number): string {
  let cursor = start;
  while (cursor > 0) {
    while (cursor > 0 && /\s/.test(source[cursor - 1])) {
      cursor -= 1;
    }
    if (cursor < 2 || source.slice(cursor - 2, cursor) !== "*/") {
      return "";
    }
    const open = source.lastIndexOf("/*", cursor - 2);
    if (open < 0) {
      return "";
    }
    const raw = source.slice(open, cursor);
    if (raw.startsWith("/**")) {
      return normalizeDoc(raw);
    }
    cursor = open;
  }
  return "";
}

function removeComments(source: string): string {
  return source.replace(/\/\*[\s\S]*?\*\//g, " ").replace(/\/\/[^\r\n]*/g, " ");
}

function normalizeSignature(source: string): string {
  return removeComments(source).replace(/\s+/g, " ").trim();
}

function findTopLevelDeclarationEnd(tokens: Token[], startIndex: number): number {
  let braceDepth = 0;
  let parenDepth = 0;
  let bracketDepth = 0;
  for (let index = startIndex; index < tokens.length; index += 1) {
    const value = tokens[index].value;
    if (value === "{") {
      braceDepth += 1;
    } else if (value === "}") {
      braceDepth -= 1;
    } else if (value === "(") {
      parenDepth += 1;
    } else if (value === ")") {
      parenDepth -= 1;
    } else if (value === "[") {
      bracketDepth += 1;
    } else if (value === "]") {
      bracketDepth -= 1;
    } else if (value === ";" && braceDepth === 0 && parenDepth === 0 && bracketDepth === 0) {
      return tokens[index].end;
    }
  }
  throw new Error(
    `unterminated exported type declaration near offset ${tokens[startIndex]?.start}`,
  );
}

function referencesInText(text: string, declarationNames: ReadonlySet<string>): string[] {
  const references: string[] = [];
  const seen = new Set<string>();
  for (const token of scanTokens(removeComments(text))) {
    if (declarationNames.has(token.value) && !seen.has(token.value)) {
      references.push(token.value);
      seen.add(token.value);
    }
  }
  return references;
}

export function parseExportedDeclarations(
  source: string,
  tokens = scanTokens(source),
): Map<string, ParsedDeclaration> {
  const declarations = new Map<string, ParsedDeclaration>();
  let braceDepth = 0;
  let parenDepth = 0;
  let bracketDepth = 0;
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token.value === "{" || token.value === "(" || token.value === "[") {
      if (token.value === "{") {
        braceDepth += 1;
      }
      if (token.value === "(") {
        parenDepth += 1;
      }
      if (token.value === "[") {
        bracketDepth += 1;
      }
      continue;
    }
    if (token.value === "}" || token.value === ")" || token.value === "]") {
      if (token.value === "}") {
        braceDepth -= 1;
      }
      if (token.value === ")") {
        parenDepth -= 1;
      }
      if (token.value === "]") {
        bracketDepth -= 1;
      }
      continue;
    }
    if (token.value !== "export" || braceDepth !== 0 || parenDepth !== 0 || bracketDepth !== 0) {
      continue;
    }
    let kindIndex = index + 1;
    if (tokens[kindIndex]?.value === "declare") {
      kindIndex += 1;
    }
    const kind = tokens[kindIndex]?.value;
    if (kind !== "interface" && kind !== "enum" && kind !== "type") {
      continue;
    }
    const nameToken = tokens[kindIndex + 1];
    if (!(nameToken && /^[A-Za-z_$]/.test(nameToken.value))) {
      throw new Error(`exported ${kind} declaration has no name near offset ${token.start}`);
    }

    let bodyStart = kindIndex + 2;
    let end: number;
    if (kind === "type") {
      while (tokens[bodyStart] && tokens[bodyStart].value !== "=") {
        bodyStart += 1;
      }
      if (!tokens[bodyStart]) {
        throw new Error(`exported type ${nameToken.value} has no equals sign`);
      }
      end = findTopLevelDeclarationEnd(tokens, bodyStart + 1);
    } else {
      while (tokens[bodyStart] && tokens[bodyStart].value !== "{") {
        bodyStart += 1;
      }
      if (!tokens[bodyStart]) {
        throw new Error(`exported ${kind} ${nameToken.value} has no body`);
      }
      const close = matchingDelimiter(tokens, bodyStart);
      end = tokens[close].end;
      if (tokens[close + 1]?.value === ";") {
        end = tokens[close + 1].end;
      }
    }

    const declaration = normalizeSignature(source.slice(token.start, end));
    declarations.set(nameToken.value, {
      name: nameToken.value,
      kind,
      description: leadingDoc(source, token.start),
      declaration,
      references: [],
    });
  }
  const names = new Set(declarations.keys());
  for (const declaration of declarations.values()) {
    declaration.references = referencesInText(declaration.declaration, names).filter(
      (name) => name !== declaration.name,
    );
  }
  return declarations;
}

/** The members of the exported interface `name`, or `undefined` when the file declares none. */
export function parseInterface(
  source: string,
  name: string,
  declarations: ReadonlyMap<string, ParsedDeclaration>,
  tokens = scanTokens(source),
): ParsedInterface | undefined {
  const start = tokens.findIndex(
    (token, index) =>
      token.value === "interface" &&
      tokens[index - 1]?.value === "export" &&
      tokens[index + 1]?.value === name,
  );
  if (start < 0) {
    return undefined;
  }
  let open = start + 2;
  const heritage: string[] = [];
  while (tokens[open] && tokens[open].value !== "{") {
    if (tokens[open - 1]?.value === "extends" || tokens[open - 1]?.value === ",") {
      heritage.push(tokens[open].value);
    }
    open += 1;
  }
  if (!tokens[open]) {
    throw new Error(`exported interface ${name} has no body`);
  }
  const close = matchingDelimiter(tokens, open);
  const declarationNames = new Set(declarations.keys());
  const parsed: ParsedInterface = {
    name,
    extends: heritage,
    methods: new Map(),
    properties: new Map(),
  };
  let index = open + 1;
  while (index < close) {
    const first = index;
    let depth = 0;
    while (index < close && !(tokens[index].value === ";" && depth === 0)) {
      const value = tokens[index].value;
      if (value === "{" || value === "(" || value === "[") {
        depth += 1;
      } else if (value === "}" || value === ")" || value === "]") {
        depth -= 1;
      }
      index += 1;
    }
    const member = tokens.slice(first, index);
    index += 1;
    let at = member[0]?.value === "readonly" ? 1 : 0;
    const nameToken = member[at];
    if (!(nameToken && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(nameToken.value))) {
      continue;
    }
    at += member[at + 1]?.value === "?" ? 2 : 1;
    const end = member[member.length - 1].end;
    const text = normalizeSignature(source.slice(nameToken.start, end));
    if (member[at]?.value === ":") {
      parsed.properties.set(
        nameToken.value,
        normalizeSignature(source.slice(member[at + 1].start, end)),
      );
      continue;
    }
    if (member[at]?.value !== "(") {
      continue;
    }
    const closeParen = matchingDelimiter(member, at);
    let arity = closeParen > at + 1 ? 1 : 0;
    for (let depth = 0, i = at + 1; i < closeParen; i += 1) {
      const value = member[i].value;
      if (value === "{" || value === "(" || value === "[" || value === "<") {
        depth += 1;
      } else if (value === "}" || value === ")" || value === "]" || value === ">") {
        depth -= 1;
      } else if (value === "," && depth === 0) {
        arity += 1;
      }
    }
    const returns = normalizeSignature(source.slice(member[closeParen + 2]?.start ?? end, end));
    const existing = parsed.methods.get(nameToken.value);
    if (existing) {
      existing.signature = `${existing.signature}\n${text}`;
      existing.references = [
        ...new Set([...existing.references, ...referencesInText(text, declarationNames)]),
      ];
      continue;
    }
    parsed.methods.set(nameToken.value, {
      name: nameToken.value,
      signature: text,
      description: leadingDoc(source, nameToken.start) || `Call ${name}.${nameToken.value}.`,
      references: referencesInText(text, declarationNames),
      returns,
      arity,
    });
  }
  return parsed;
}
