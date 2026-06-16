/**
 * 微信 JS-SDK 签名接口
 * GET /api/wechat-sign?url=<当前页面URL>
 * 返回 { appId, timestamp, nonceStr, signature }
 */

// 内存缓存（Worker 实例级别，适合低流量场景）
let _token = null, _tokenExp = 0;
let _ticket = null, _ticketExp = 0;

async function fetchToken(appId, appSecret) {
  const now = Date.now();
  if (_token && now < _tokenExp) return _token;

  const res = await fetch(
    `https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid=${appId}&secret=${appSecret}`
  );
  const data = await res.json();
  if (data.access_token) {
    _token = data.access_token;
    _tokenExp = now + (data.expires_in - 300) * 1000; // 提前 5 分钟过期
    return _token;
  }
  throw new Error('获取 access_token 失败: ' + JSON.stringify(data));
}

async function fetchTicket(token) {
  const now = Date.now();
  if (_ticket && now < _ticketExp) return _ticket;

  const res = await fetch(
    `https://api.weixin.qq.com/cgi-bin/ticket/getticket?access_token=${token}&type=jsapi`
  );
  const data = await res.json();
  if (data.errcode === 0) {
    _ticket = data.ticket;
    _ticketExp = now + (data.expires_in - 300) * 1000;
    return _ticket;
  }
  throw new Error('获取 jsapi_ticket 失败: ' + JSON.stringify(data));
}

function nonceStr() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let s = '';
  for (let i = 0; i < 16; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

async function sha1(str) {
  const enc = new TextEncoder().encode(str);
  const hash = await crypto.subtle.digest('SHA-1', enc);
  return Array.from(new Uint8Array(hash))
    .map(b => b.toString(16).padStart(2, '0'))
    .join('');
}

export async function onRequest(context) {
  const { request, env } = context;
  const cors = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
  };

  if (request.method === 'OPTIONS') {
    return new Response(null, { status: 204, headers: cors });
  }

  const url = new URL(request.url);
  const signUrl = url.searchParams.get('url') || '';

  const appId = env.WX_APP_ID || 'wx52f55029c4340a68';
  const appSecret = env.WX_APP_SECRET;

  if (!appSecret) {
    return new Response(JSON.stringify({ error: 'WX_APP_SECRET 未配置' }), {
      status: 500, headers: cors
    });
  }

  try {
    const token = await fetchToken(appId, appSecret);
    const ticket = await fetchTicket(token);

    const n = nonceStr();
    const ts = Math.floor(Date.now() / 1000);
    const raw = `jsapi_ticket=${ticket}&noncestr=${n}&timestamp=${ts}&url=${signUrl}`;
    const sig = await sha1(raw);

    return new Response(JSON.stringify({
      appId,
      timestamp: ts,
      nonceStr: n,
      signature: sig,
    }), { headers: cors });
  } catch (e) {
    return new Response(JSON.stringify({ error: e.message }), {
      status: 500, headers: cors
    });
  }
}
