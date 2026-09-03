const http = require('http');
const { WebSocketServer } = require('ws');

const port = parsePositiveInteger(process.env.PORT, 8080);
const tickIntervalMs = parsePositiveInteger(process.env.TICK_INTERVAL_MS, 1000);
const symbolsBySocket = new Map();
const marketData = new Map();

const server = http.createServer();
const webSocketServer = new WebSocketServer({ server });

webSocketServer.on('connection', (socket) => {
  symbolsBySocket.set(socket, new Set());

  socket.on('message', (message) => {
    handleMessage(socket, message);
  });

  socket.on('close', () => {
    symbolsBySocket.delete(socket);
  });
});

setInterval(() => {
  for (const [socket, symbols] of symbolsBySocket) {
    if (socket.readyState !== socket.OPEN || symbols.size === 0) {
      continue;
    }

    const changedData = {};
    for (const symbol of symbols) {
      changedData[symbol] = advanceMarketData(getMarketData(symbol));
    }

    socket.send(JSON.stringify({ type: 'tick', data: changedData }));
  }
}, tickIntervalMs);

server.listen(port, () => {
  console.log(`Mock server listening on ws://localhost:${port}`);
});

function handleMessage(socket, message) {
  let request;
  try {
    request = JSON.parse(message.toString());
  } catch {
    sendError(socket, 'INVALID_REQUEST', 'message must be valid JSON');
    return;
  }

  if (!request || !['subscribe', 'unsubscribe'].includes(request.action)) {
    sendError(socket, 'INVALID_REQUEST', 'action must be subscribe or unsubscribe');
    return;
  }

  if (!Array.isArray(request.symbols) || request.symbols.length === 0) {
    sendError(socket, 'INVALID_REQUEST', 'symbols must be a non-empty array');
    return;
  }

  if (!request.symbols.every(isValidSymbol)) {
    sendError(socket, 'INVALID_REQUEST', 'symbols must contain only letters, digits, ., and -');
    return;
  }

  const symbols = symbolsBySocket.get(socket);
  if (!symbols) {
    return;
  }

  if (request.action === 'subscribe') {
    const snapshot = {};
    for (const symbol of request.symbols) {
      symbols.add(symbol);
      snapshot[symbol] = getMarketData(symbol);
    }
    socket.send(JSON.stringify({ type: 'snapshot', data: snapshot }));
  } else {
    for (const symbol of request.symbols) {
      symbols.delete(symbol);
    }
  }
}

function getMarketData(symbol) {
  if (!marketData.has(symbol)) {
    const open = deterministicPrice(symbol);
    marketData.set(symbol, {
      LAST: open,
      OPEN: open,
      HIGH: open,
      LOW: open,
      VOLUME: 100000,
    });
  }

  return { ...marketData.get(symbol) };
}

function advanceMarketData(data) {
  const previousLast = data.LAST;
  const nextLast = roundPrice(previousLast * (1 + (Math.random() - 0.5) * 0.01));
  data.LAST = nextLast;
  data.VOLUME += Math.floor(Math.random() * 1000) + 1;

  const changed = { LAST: data.LAST, VOLUME: data.VOLUME };
  if (nextLast > data.HIGH) {
    data.HIGH = nextLast;
    changed.HIGH = nextLast;
  }
  if (nextLast < data.LOW) {
    data.LOW = nextLast;
    changed.LOW = nextLast;
  }

  return changed;
}

function deterministicPrice(symbol) {
  let hash = 0;
  for (const character of symbol) {
    hash = (hash * 31 + character.charCodeAt(0)) >>> 0;
  }
  return roundPrice(50 + (hash % 1950));
}

function roundPrice(value) {
  return Math.round(value * 100) / 100;
}

function isValidSymbol(symbol) {
  return typeof symbol === 'string' && /^[A-Za-z0-9.-]+$/.test(symbol);
}

function parsePositiveInteger(value, fallback) {
  const parsed = Number.parseInt(value, 10);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

function sendError(socket, code, message) {
  socket.send(JSON.stringify({ type: 'error', code, message }));
}
