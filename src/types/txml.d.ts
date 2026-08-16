declare module 'txml/txml' {
  export interface TxmlNode {
    tagName: string;
    attributes: Record<string, string>;
    children: Array<TxmlNode | string>;
  }

  export interface TxmlParseOptions {
    keepComments?: boolean;
    keepWhitespace?: boolean;
    simplify?: boolean;
  }

  export function parse(
    source: string,
    options?: TxmlParseOptions,
  ): Array<TxmlNode | string>;
}
