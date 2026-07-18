const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders })

  try {
    const apiKey = Deno.env.get('WEATHERBIT_API_KEY')
    if (!apiKey) return new Response(JSON.stringify({ error: 'WEATHERBIT_API_KEY is not configured' }), { status: 503, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })

    const incoming = new URL(request.url)
    const params = new URLSearchParams({
      lat: incoming.searchParams.get('lat') || '52.69',
      lon: incoming.searchParams.get('lon') || '0.95',
      search_distance_km: incoming.searchParams.get('search_distance_km') || '25',
      search_mins: incoming.searchParams.get('search_mins') || '45',
      limit: '50',
      sort: incoming.searchParams.get('sort') || 'distance',
      key: apiKey,
    })
    const upstream = await fetch(`https://api.weatherbit.io/v2.0/current/lightning?${params}`)
    if (upstream.status === 204) return new Response(JSON.stringify({ lightning: [] }), { status: 200, headers: { ...corsHeaders, 'Content-Type': 'application/json', 'Cache-Control': 'public, max-age=240' } })
    const body = await upstream.text()
    return new Response(body, { status: upstream.status, headers: { ...corsHeaders, 'Content-Type': 'application/json', 'Cache-Control': 'public, max-age=240' } })
  } catch (error) {
    return new Response(JSON.stringify({ error: error instanceof Error ? error.message : 'Lightning service failed' }), { status: 500, headers: { ...corsHeaders, 'Content-Type': 'application/json' } })
  }
})
