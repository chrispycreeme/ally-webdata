import { initializeApp } from "https://www.gstatic.com/firebasejs/12.3.0/firebase-app.js";
import { getAnalytics } from "https://www.gstatic.com/firebasejs/12.3.0/firebase-analytics.js";
// TODO: Add SDKs for Firebase products that you want to use
// https://firebase.google.com/docs/web/setup#available-libraries

// Your web app's Firebase configuration
// For Firebase JS SDK v7.20.0 and later, measurementId is optional
import {
    getFirestore, collection, addDoc, getDocs, doc, GeoPoint, setDoc
} from "https://www.gstatic.com/firebasejs/12.3.0/firebase-firestore.js";

import {
    getAuth, signInAnonymously, onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/12.3.0/firebase-auth.js";

const firebaseConfig = {
    apiKey: "AIzaSyA5pyu0LlfvW06m1jdwVXVW8JlW5G7eXps",
    authDomain: "ally-database-669a4.firebaseapp.com",
    projectId: "ally-database-669a4",
    storageBucket: "ally-database-669a4.firebasestorage.app",
    messagingSenderId: "954693497275",
    appId: "1:954693497275:web:16cf70e20465170949149e",
    measurementId: "G-1RGKL7S7X7"
};

const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);
const auth = getAuth(app);

const db = getFirestore(app);

const form = document.getElementById('buildingForm');
const buildings = document.getElementById('buildingList');
const refreshList = document.getElementById('refreshList');
// ADD: map elements
const mapContainer = document.getElementById('buildingMap');
let map;
let polygonLayer;

function ensureMap() {
    if (!mapContainer || map) return;
    map = L.map('buildingMap', { preferCanvas: true }).setView([0, 0], 2);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; OpenStreetMap contributors'
    }).addTo(map);
    polygonLayer = L.layerGroup().addTo(map);
}

function parseCoords(raw) {
    const [latStr, lonStr] = raw.split(',').map(s => s.trim());
    const lat = parseFloat(latStr);
    const lon = parseFloat(lonStr);
    if (Number.isNaN(lat) || Number.isNaN(lon)) {
        throw new Error(`Invalid coordinate: ${raw}`);
    }
    return new GeoPoint(lat, lon);
}

// REPLACE loadBuildings with map-enabled version
async function loadBuildings() {
    ensureMap();
    buildings.innerHTML = 'Loading...';
    const snapshot = await getDocs(collection(db, 'buildings'));
    buildings.innerHTML = '';
    if (polygonLayer) polygonLayer.clearLayers();
    const boundsList = [];

    snapshot.forEach(d => {
        const data = d.data();
        const li = document.createElement('li');
        li.textContent = `${d.id} (Level ${data.level}) - Color: ${data.color}`;
        buildings.appendChild(li);

        const pts = Array.isArray(data.polygons)
            ? data.polygons
                .map(p => (p && typeof p.latitude === 'number' && typeof p.longitude === 'number'
                    ? [p.latitude, p.longitude]
                    : null))
                .filter(Boolean)
            : [];

        if (pts.length >= 3 && polygonLayer) {
            const poly = L.polygon(pts, {
                color: data.color || '#7c3aed',
                weight: 3,
                fillOpacity: 0.25
            }).addTo(polygonLayer);
            boundsList.push(poly.getBounds());
        }
    });

    if (!buildings.children.length) {
        buildings.innerHTML = '<i>No buildings found</i>';
    } else if (map && boundsList.length) {
        const combined = boundsList.reduce((acc, b) => acc.extend(b), boundsList[0]);
        map.fitBounds(combined, { padding: [24, 24] });
    }
}

form.addEventListener('submit', async (e) => {
    e.preventDefault();
    try {
        const name = document.getElementById('name').value.trim();
        let color = document.getElementById('color').value.trim();
        const level = document.getElementById('level').value;

        const pointValues = ['pointA', 'pointB', 'pointC', 'pointD']
            .map(id => document.getElementById(id).value.trim())
            .filter(v => v.length);

        if (pointValues.length < 3) {
            throw new Error('At least 3 points are required to form a polygon');
        }

        const polygons = pointValues.map(parseCoords);

        if (!name) throw new Error('Name needed');
        if (!/^#?[0-9a-f]{6}$/i.test(color)) throw new Error('Invalid color hex');
        if (!color.startsWith('#')) color = '#' + color;
        const levelParsed = parseInt(level, 10);
        if (Number.isNaN(levelParsed)) throw new Error('Invalid level');

        await setDoc(doc(db, 'buildings', name), {
            color,
            level: levelParsed,
            polygons
        });

        form.reset();
        await loadBuildings();
    } catch (err) {
        alert(err.message);
    }
});

refreshList.addEventListener('click', () => {
    if (auth.currentUser) loadBuildings();
});
signInAnonymously(auth).catch(err => {
    console.error('Anon auth failed', err);
});
onAuthStateChanged(auth, user => {
    if (user) {
        ensureMap();
        loadBuildings();
    }
});