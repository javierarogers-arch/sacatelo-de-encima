const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz36AA4zd3nmjbCNCBRSU_fuCs-EJONdY6HNJDHF03HUt4xFxEuLY_si5wI_fwwzeU8/exec';

exports.handler = async function(event) {
  try {
    const params = new URLSearchParams(event.queryStringParameters || {});
    if (!params.get('action')) params.set('action', 'productos');
    const url = `${APPS_SCRIPT_URL}?${params.toString()}`;
    const response = await fetch(url, { redirect: 'follow' });
    const body = await response.text();

    return {
      statusCode: response.status,
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store'
      },
      body
    };
  } catch (error) {
    return {
      statusCode: 500,
      headers: { 'Content-Type': 'application/json; charset=utf-8' },
      body: JSON.stringify({ ok: false, error: error.message })
    };
  }
};
