async function getMapUrl(city, state, country) {

    const cacheKey = `Map_Coord_${city},${state},${country}`;

    let cachedPoint = window.localStorage.getItem(cacheKey);
    let point = { x: 0, y: 0};      // Assume nobody lives at this spot in the Atlantic

    if (cachedPoint) {

        point = JSON.parse(cachedPoint);

    } else {

        const response = await fetch(`http://dev.virtualearth.net/REST/v1/Locations?countryRegion=${country}&adminDistrict=${state}&locality={city}&key=${constants.bingMapsApiKey}`,
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

    if (point.x || point.y) {
        return `http://dev.virtualearth.net/REST/v1/Imagery/Map/Aerial/0,0/0?mapSize=500,300&pp=${point.x},${point.y};22&key=${constants.bingMapsApiKey}`;
    } else {
        return "#";
    }
}
