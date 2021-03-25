
export async function getMapUrl(places) {

    let points = "";
    let pointCount = 0;
    places.forEach(async (place) => {

        const { city, state, country } = place;
        console.log(`CITY=${city} STATE=${state} COUNTRY=${country}`);
        const cacheKey = `Map_Coord_${city},${state},${country}`;

        let cachedPoint = window.localStorage.getItem(cacheKey);
        let point = { x: 0, y: 0 };      // Assume nobody lives at this spot in the Atlantic

        if (cachedPoint) {

            point = JSON.parse(cachedPoint);

        } else {

            const response = await fetch(`http://dev.virtualearth.net/REST/v1/Locations?countryRegion=${country}&adminDistrict=${state}&locality=${city}&key=${constants.bingMapsApiKey}`,
                {
                    method: 'GET',
                    headers: { "accept": "application/json" },
                });
                
            if (response.ok) {
                const responseJson = await response.json();
                const geoLocation = responseJson;
                point = {
                    x: geoLocation.resourceSets[0].resources[0].point.coordinates[0],
                    y: geoLocation.resourceSets[0].resources[0].point.coordinates[1]
                };
            } else {
                console.log(`Error getting map location ${response.status}: ${response.statusText}`);
            }
            
            window.localStorage.setItem(cacheKey, JSON.stringify(point));
        }
        // If we got a point and if there are less than 18 points, add it to the map
        if ((point.x || point.y) && pointCount++ < 18) {
            points += `&pp=${point.x},${point.y};22`;
        }
    });
    
    console.log(`http://dev.virtualearth.net/REST/v1/Imagery/Map/Aerial/0,0/1?mapSize=500,300${points}&key=${constants.bingMapsApiKey}`);

    return `http://dev.virtualearth.net/REST/v1/Imagery/Map/Aerial/0,0/1?mapSize=500,300${points}&key=${constants.bingMapsApiKey}`;
}

export async function getTimezoneInfo(city, state, country) {
    const query = [city, state, country].join(', ');
    const result = await fetch(`https://dev.virtualearth.net/REST/v1/TimeZone/?query=${query}&key=${constants.bingMapsApiKey}`);
    if (result.ok) {
        const json = await result.json();
        return json.resourceSets[0].resources[0].timeZoneAtLocation[0].timeZone[0];
    }
    else {
        throw result.error;
    }
}

