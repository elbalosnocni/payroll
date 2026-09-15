function now_() {
  return new Date();
}


function formatDateTime_(date) {
  const value = date instanceof Date ? date : new Date(date);

  if (isNaN(value.getTime())) {
    return '';
  }

  return Utilities.formatDate(
    value,
    Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh',
    'dd/MM/yyyy HH:mm:ss'
  );
}


function formatDateTimeFromEpoch_(epochMilliseconds) {
  return formatDateTime_(new Date(Number(epochMilliseconds)));
}


function normalizeString_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  return String(value).trim();
}


function normalizeEmployeeCode_(value) {
  return normalizeString_(value).toUpperCase();
}


function normalizeFullName_(value) {
  return normalizeString_(value)
    .replace(/\s+/g, ' ')
    .toLocaleLowerCase('vi-VN');
}


function normalizeCCCD_(value) {
  if (value === null || value === undefined) {
    return '';
  }

  let result = String(value).trim();

  if (/^\d+\.0+$/.test(result)) {
    result = result.replace(/\.0+$/, '');
  }

  return result;
}


function isValidCCCD_(value) {
  return /^\d{9,12}$/.test(normalizeCCCD_(value));
}


function bytesToHex_(bytes) {
  return bytes
    .map(function(byte) {
      const value = byte < 0 ? byte + 256 : byte;
      return ('0' + value.toString(16)).slice(-2);
    })
    .join('');
}


function hexToBytes_(hex) {
  const result = [];

  for (let index = 0; index < hex.length; index += 2) {
    result.push(parseInt(hex.substr(index, 2), 16));
  }

  return result;
}


function randomHex_(byteCount) {
  return bytesToHex_(
    Utilities
      .newBlob(
        Utilities.getUuid()
      )
      .getBytes()
      .concat(
        Utilities.getUuid().split('').map(function(character) {
          return character.charCodeAt(0);
        })
      )
      .slice(0, byteCount)
  );
}


function randomToken_(byteCount) {
  const seed = [
    Utilities.getUuid(),
    Utilities.getUuid(),
    String(new Date().getTime()),
    Math.random().toString(36)
  ].join('|');

  let bytes = Utilities.newBlob(seed).getBytes();

  while (bytes.length < byteCount) {
    bytes = bytes.concat(
      Utilities
        .computeDigest(
          Utilities.DigestAlgorithm.SHA_256,
          Utilities.getUuid()
        )
    );
  }

  return Utilities.base64EncodeWebSafe(
    bytes.slice(0, byteCount)
  ).replace(/=+$/g, '');
}


function sha256Bytes_(bytes) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    bytes
  );
}


function sha256Hex_(text) {
  return bytesToHex_(
    sha256Bytes_(
      Utilities.newBlob(
        String(text),
        'text/plain',
        'data'
      ).getBytes()
    )
  );
}


function hashPassword_(password, saltHex) {
  const passwordValue = String(password);
  const salt = hexToBytes_(saltHex);

  let input = salt.concat(
    Utilities.newBlob(passwordValue).getBytes()
  );

  let digest = sha256Bytes_(input);

  for (
    let iteration = 1;
    iteration < CONFIG.PASSWORD.ITERATIONS;
    iteration++
  ) {
    digest = sha256Bytes_(
      digest.concat(
        Utilities.newBlob(passwordValue).getBytes()
      )
    );
  }

  return bytesToHex_(digest);
}


function createPasswordRecord_(password) {
  const saltHex = randomHex_(CONFIG.PASSWORD.SALT_BYTES);

  return {
    passwordHash: hashPassword_(password, saltHex),
    passwordSalt: saltHex
  };
}


function timingSafeEqual_(left, right) {
  const a = String(left || '');
  const b = String(right || '');

  let result = a.length ^ b.length;
  const maxLength = Math.max(a.length, b.length);

  for (let index = 0; index < maxLength; index++) {
    result |= (
      (index < a.length ? a.charCodeAt(index) : 0) ^
      (index < b.length ? b.charCodeAt(index) : 0)
    );
  }

  return result === 0;
}


function signToken_(payload) {
  const encodedPayload = Utilities.base64EncodeWebSafe(
    Utilities.newBlob(
      JSON.stringify(payload)
    ).getBytes()
  ).replace(/=+$/g, '');

  const signature = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(
      encodedPayload,
      getSecretKey_()
    )
  ).replace(/=+$/g, '');

  return encodedPayload + '.' + signature;
}


function verifyToken_(token) {
  try {
    const parts = String(token || '').split('.');

    if (parts.length !== 2) {
      return null;
    }

    const payloadEncoded = parts[0];
    const suppliedSignature = parts[1];

    const expectedSignature = Utilities.base64EncodeWebSafe(
      Utilities.computeHmacSha256Signature(
        payloadEncoded,
        getSecretKey_()
      )
    ).replace(/=+$/g, '');

    if (!timingSafeEqual_(
      suppliedSignature,
      expectedSignature
    )) {
      return null;
    }

    const decoded = Utilities.base64DecodeWebSafe(
      payloadEncoded
    );

    const payload = JSON.parse(
      Utilities.newBlob(decoded).getDataAsString('UTF-8')
    );

    if (!payload.exp || Number(payload.exp) < Date.now()) {
      return null;
    }

    return payload;
  } catch (error) {
    return null;
  }
}


function createUserToken_(employeeCode, mustChangePassword) {
  const issuedAt = Date.now();
  const expiresAt =
    issuedAt +
    CONFIG.API.TOKEN_TTL_SECONDS * 1000;

  return signToken_({
    typ: 'USER',
    sub: normalizeEmployeeCode_(employeeCode),
    iat: issuedAt,
    exp: expiresAt,
    mcp: Boolean(mustChangePassword),
    jti: randomToken_(CONFIG.SESSION.TOKEN_BYTES)
  });
}


function createAdminToken_(adminId, username) {
  const issuedAt = Date.now();
  const expiresAt =
    issuedAt +
    CONFIG.API.TOKEN_TTL_SECONDS * 1000;

  return signToken_({
    typ: 'ADMIN',
    sub: String(adminId),
    username: String(username),
    iat: issuedAt,
    exp: expiresAt,
    jti: randomToken_(CONFIG.SESSION.TOKEN_BYTES)
  });
}


function requireUserToken_(token) {
  const payload = verifyToken_(token);

  if (!payload || payload.typ !== 'USER') {
    throw new Error('UNAUTHORIZED');
  }

  return payload;
}


function requireAdminToken_(token) {
  const payload = verifyToken_(token);

  if (!payload || payload.typ !== 'ADMIN') {
    throw new Error('ADMIN_UNAUTHORIZED');
  }

  return payload;
}


function getRequestClientIp_(e) {
  try {
    if (
      e &&
      e.parameter &&
      e.parameter.ip
    ) {
      return String(e.parameter.ip).substring(0, 64);
    }
  } catch (error) {
  }

  return '';
}


function safeJsonParse_(value) {
  if (value === null || value === undefined || value === '') {
    return {};
  }

  if (typeof value === 'object') {
    return value;
  }

  return JSON.parse(String(value));
}


function isObject_(value) {
  return (
    value !== null &&
    typeof value === 'object' &&
    !Array.isArray(value)
  );
}


function assertString_(value, fieldName, minLength, maxLength) {
  const text = String(value || '').trim();

  if (
    text.length < minLength ||
    text.length > maxLength
  ) {
    throw new Error(
      'INVALID_' + fieldName.toUpperCase()
    );
  }

  return text;
}


function validatePassword_(password) {
  const value = String(password || '');

  if (
    value.length < CONFIG.PASSWORD.MIN_LENGTH ||
    value.length > CONFIG.PASSWORD.MAX_LENGTH
  ) {
    throw new Error('INVALID_PASSWORD_LENGTH');
  }

  return value;
}


function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });
}


function rowsToObjects_(headers, rows) {
  return rows.map(function(row) {
    const object = {};

    headers.forEach(function(header, index) {
      object[header] = row[index];
    });

    return object;
  });
}


function objectToRow_(headers, object) {
  return headers.map(function(header) {
    return object[header] !== undefined
      ? object[header]
      : '';
  });
}


function findColumnIndex_(headers, columnName) {
  const index = headers.indexOf(columnName);

  if (index < 0) {
    throw new Error(
      'MISSING_COLUMN_' + columnName
    );
  }

  return index;
}


function createResponse_(success, data, message, code) {
  return {
    success: Boolean(success),
    code: code || (success ? 'OK' : 'ERROR'),
    message: message || '',
    data: data === undefined ? null : data,
    serverTime: formatDateTime_(new Date())
  };
}


function auditLog_(
  actorType,
  actorId,
  action,
  target,
  success,
  ipAddress,
  details
) {
  try {
    const sheet = getSheet_(CONFIG.SHEETS.AUDIT);

    sheet.appendRow([
      formatDateTime_(new Date()),
      String(actorType || ''),
      String(actorId || ''),
      String(action || ''),
      String(target || ''),
      Boolean(success),
      String(ipAddress || ''),
      String(details || '').substring(0, 5000)
    ]);
  } catch (error) {
    console.error('auditLog_ failed: ' + error.message);
  }
}


function acquireScriptLock_(timeoutMilliseconds) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(timeoutMilliseconds || 10000)) {
    throw new Error('SERVER_BUSY');
  }

  return lock;
}
