/**
 * Connection protocol for recording client's unique identity on the server side.
 * Since the underlying communication is based on penpal, there's no need for TCP-like
 * three-way handshake to confirm message delivery. The connection is established
 * as soon as the server receives the client's identity information.
 * The uniqueness of client identity is guaranteed by the client itself, while the server
 * is only responsible for recording it. Since the client is essentially a page embedded
 * in the Office SDK iframe, its identity uniqueness is easily guaranteed.
 *
 * The connection process involves three scenarios:
 *
 * 1. Client-initiated connection
 * In this case, the server usually exists but might not be fully initialized.
 * If the server is not ready, the client needs to wait until the server is ready
 * before sending the confirmation request. The client actively provides its identity
 * information, and the server records it and returns a response.
 *
 * 2. Server-initiated connection
 * Here, the server is ready, but the client might not exist (page not loaded in iframe
 * or restricted by same-origin policy). The server requests the client's identity
 * information, and the client responds to complete the connection. The server may
 * receive identity information from 0 to multiple clients.
 *
 * 3. Simultaneous connection attempts
 * Since connection establishment is a bidirectional asynchronous process, both sides
 * might send connection requests simultaneously before the connection is complete.
 * This scenario requires no special handling as the server automatically deduplicates.
 */

/**
 * Server protocol interface
 * These interfaces are for remote invocation by clients, not for server's own use
 */
export type ServerProtocol = {
  /**
   * Client registers its identity with the server
   * All subsequent calls must include this clientId as identity token
   * @param clientId Client identity token
   * @returns Whether the connection was established successfully
   */
  open: (clientId: string) => boolean;
  /**
   * Server actively closes the client connection
   * @param clientId Client identity token
   * @returns Whether the connection was closed successfully
   */
  close: (clientId: string) => boolean;
};

/**
 * Context required for server initialization
 */
interface ServerContext {
  addClient: (id: string) => void;
  deleteClient: (id: string) => void;
}

export function createServerProtocol(context: ServerContext): ServerProtocol {
  return {
    open: (clientId: string): boolean => {
      // TODO: Should throw error on duplicate registration
      context.addClient(clientId);

      return true;
    },

    close: (clientId: string): boolean => {
      // TODO: Should throw error if clientId doesn't exist
      context.deleteClient(clientId);

      return true;
    },
  };
}

/**
 * Client protocol interface
 * These interfaces are for remote invocation by server, not for client's own use
 */
export type ClientProtocol = {
  open: () => string[];
  close: (clientId: string) => void;
};

/**
 * Context required for client initialization
 */
interface ClientContext {
  /**
   * Get identity information of connected clients
   */
  getClients: () => Set<string>;
}

export function createClientProtocol(context: ClientContext): ClientProtocol {
  return {
    open: (): string[] => {
      return Array.from(context.getClients());
    },
    close: (clientId: string): void => {
      // TODO: Record the incoming clientId
    },
  };
}

export const OfficeSdkRpcChannel = '#office-sdk-rpc';

export enum RPCArgType {
  String = 'string',
  Number = 'number',
  Boolean = 'boolean',
  Object = 'object',
  Callback = 'callback',
  Null = 'null',
  Undefined = 'undefined',
  Array = 'array',
}

/**
 * 可通过 penpal 调用的参数形式我们约束为合法的 JSON 类型
 */
export type RPCParameter =
  | string
  | number
  | boolean
  | RPCMethod
  | null
  | undefined
  | RPCParameter[]
  | { [key: string]: RPCParameter };

/**
 * 可通过 penpal 调用的方法，返回值约束为合法的 JSON 类型
 */
export type RPCMethod = (...args: any[]) => any;

/**
 *
 */
export type RPCMethods = {
  [index: string]: RPCMethod;
};

/**
 * 远程调用协议，用于定义服务端提供的远程调用方法。
 * 服务端需要基于这个协议提供 penpal 的 Methods 实现，
 * 客户端基于这个协议可以创建 penpal 的 RemoteProxy 实现
 */
export type RPCProtocol<TMethods extends RPCMethods> = {
  client: {
    [K in keyof TMethods]: TMethods[K];
  };
  server: TMethods;
};
