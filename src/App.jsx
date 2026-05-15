import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Upload, Play, Pause, Download, Music, Trash2, Clock, Repeat, FileAudio, ChevronRight, Settings, Youtube, Key, Mic, Loader2, Check, AlertCircle, Copy, Link2, ListMusic, ExternalLink, RefreshCw, FileText, Save, FolderOpen, Plus, ChevronUp, ChevronDown, GripVertical, Award } from 'lucide-react';

// ===== Default exam =====
const DEFAULT_QUESTIONS = [
  { id: 1, label: 'Extract 1', plays: 3, gapBetweenPlays: 30, gapAfter: 60, marks: 12, intro: 'Extract 1. You will hear this extract three times.', source: null },
  { id: 2, label: 'Extract 2', plays: 3, gapBetweenPlays: 30, gapAfter: 60, marks: 12, intro: 'Extract 2. You will hear this extract three times.', source: null },
  { id: 3, label: 'Extract 3', plays: 3, gapBetweenPlays: 30, gapAfter: 60, marks: 9, intro: 'Extract 3. You will hear this extract three times.', source: null },
  { id: 4, label: 'Extract 4', plays: 3, gapBetweenPlays: 30, gapAfter: 60, marks: 12, intro: 'Extract 4. You will hear this extract three times.', source: null },
  { id: 5, label: 'Extract 5', plays: 3, gapBetweenPlays: 30, gapAfter: 60, marks: 12, intro: 'Extract 5. You will hear this extract three times.', source: null },
  { id: 6, label: 'Extract 6', plays: 2, gapBetweenPlays: 20, gapAfter: 45, marks: 3, intro: 'Extract 6. You will hear this extract two times.', source: null },
  { id: 7, label: 'Extract 7', plays: 3, gapBetweenPlays: 25, gapAfter: 45, marks: 7, intro: 'Extract 7. You will hear this extract three times.', source: null },
  { id: 8, label: 'Extract 8', plays: 3, gapBetweenPlays: 25, gapAfter: 30, marks: 8, intro: 'Extract 8. You will hear this extract three times. This is the final extract.', source: null },
];

const DEFAULT_SCRIPT = {
  opening: 'This is the Music listening examination. You will now have five minutes to read through all of the listening questions. You may not write anything during this time.',
  postReading: 'Your five minutes of reading time is now over. The listening section will now begin.',
  // {n} = play number as a numeral (2, 3, 4...). {ord} = ordinal word (second, third, fourth...). {final} expands to " and final" when this is the last play, else empty.
  betweenPlays: 'You will now hear the extract for the {ord}{final} time.',
  ending: 'This is the end of the listening section of the examination.',
};

const ORDINAL_WORDS = ['', 'first', 'second', 'third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth', 'ninth', 'tenth'];

let pdfjsPromise = null;
function loadPdfJs() {
  if (pdfjsPromise) return pdfjsPromise;
  pdfjsPromise = new Promise((resolve, reject) => {
    if (window.pdfjsLib) return resolve(window.pdfjsLib);
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(window.pdfjsLib);
    };
    script.onerror = () => reject(new Error('Failed to load PDF.js'));
    document.head.appendChild(script);
  });
  return pdfjsPromise;
}

// Word-to-number converter for "three times" -> 3
const NUMBER_WORDS = {
  one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
  once: 1, twice: 2, thrice: 3,
};

function parsePlaysCount(text) {
  // Try digit first: "3 times", "hear this 3 times"
  const digitMatch = text.match(/(\d+)\s*times?/i);
  if (digitMatch) return parseInt(digitMatch[1]);
  // Word form
  for (const [word, n] of Object.entries(NUMBER_WORDS)) {
    if (new RegExp(`\\b${word}\\s*times?\\b`, 'i').test(text)) return n;
    if (new RegExp(`\\b${word}\\b`, 'i').test(text) && /times?/i.test(text)) return n;
  }
  return null;
}

// Parse a music exam PDF and return detected extracts
async function parseExamPdf(file) {
  const pdfjs = await loadPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise;

  // Collect text page-by-page so we can locate extract headings
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    // Join with spaces, but track items so we can roughly reconstruct lines
    const lines = [];
    let currentLine = '';
    let lastY = null;
    for (const item of content.items) {
      const y = item.transform ? item.transform[5] : null;
      if (lastY !== null && y !== null && Math.abs(y - lastY) > 5) {
        if (currentLine.trim()) lines.push(currentLine.trim());
        currentLine = '';
      }
      currentLine += item.str + ' ';
      lastY = y;
    }
    if (currentLine.trim()) lines.push(currentLine.trim());
    pages.push({ pageNum: i, lines, raw: lines.join('\n') });
  }

  // Combined text for global searches
  const fullText = pages.map(p => p.raw).join('\n');

  // === Extract title (best effort) ===
  let title = '';
  for (const line of pages[0]?.lines || []) {
    if (/exam|paper|listening|aural|music/i.test(line) && line.length < 100 && line.length > 4) {
      title = title ? `${title} — ${line}` : line;
      if (title.length > 60) break;
    }
  }

  // === Pull marks table (if present) ===
  // Looks for patterns like "Extract 1 12" or "Extract 1   12 marks"
  const marksMap = new Map();
  const marksTableRegex = /Extract\s+(\d+)\s+(\d+)/gi;
  let m;
  while ((m = marksTableRegex.exec(fullText)) !== null) {
    const num = parseInt(m[1]);
    const marks = parseInt(m[2]);
    if (marks > 0 && marks < 100 && !marksMap.has(num)) {
      marksMap.set(num, marks);
    }
  }

  // === Find extract headings + plays count ===
  // For each "Extract N" heading, look at the next ~200 chars for play count info
  const extractRegex = /(?:^|\n|\s)(Extract|Question)\s+(\d+)\b/gi;
  const extracts = new Map(); // num -> {label, plays, raw context}

  let match;
  while ((match = extractRegex.exec(fullText)) !== null) {
    const kind = match[1]; // "Extract" or "Question"
    const num = parseInt(match[2]);
    // Look at next 300 chars for play-count phrase
    const context = fullText.slice(match.index, match.index + 400);
    const plays = parsePlaysCount(context);
    if (!extracts.has(num) || (plays && !extracts.get(num).plays)) {
      extracts.set(num, {
        num,
        label: `${kind} ${num}`,
        plays: plays || 3, // default to 3 if not detected
        marks: marksMap.get(num) || null,
        detectedPlays: plays !== null,
      });
    }
  }

  // Sort by number
  const sorted = Array.from(extracts.values()).sort((a, b) => a.num - b.num);

  return {
    title,
    pageCount: pdf.numPages,
    extracts: sorted,
    marksTableFound: marksMap.size > 0,
    fullText, // returned for debugging if needed
  };
}

function buildIntroForExtract(label, plays, isLast = false) {
  const word = plays === 1 ? 'once' : plays === 2 ? 'two times' : plays === 3 ? 'three times' : `${plays} times`;
  return `${label}. You will hear this extract ${word}.${isLast ? ' This is the final extract.' : ''}`;
}

function renderBetweenPlays(template, playNumber, isFinal) {
  return template
    .replace(/\{n\}/g, String(playNumber))
    .replace(/\{ord\}/g, ORDINAL_WORDS[playNumber] || ordinal(playNumber))
    .replace(/\{final\}/g, isFinal ? ' and final' : '');
}

function ordinal(n) {
  const s = ['th', 'st', 'nd', 'rd'];
  const v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

function formatTime(seconds) {
  if (!isFinite(seconds)) return '0:00';
  const m = Math.floor(seconds / 60);
  const s = Math.floor(seconds % 60);
  return `${m}:${s.toString().padStart(2, '0')}`;
}

function parseTimestamp(str) {
  if (!str) return 0;
  str = String(str).trim();
  if (!str) return 0;
  const hms = str.match(/^(?:(\d+)h)?(?:(\d+)m)?(?:(\d+(?:\.\d+)?)s)?$/i);
  if (hms && (hms[1] || hms[2] || hms[3])) {
    return (parseInt(hms[1] || 0) * 3600) + (parseInt(hms[2] || 0) * 60) + parseFloat(hms[3] || 0);
  }
  const parts = str.split(':').map(p => parseFloat(p));
  if (parts.some(isNaN)) return 0;
  if (parts.length === 1) return parts[0];
  if (parts.length === 2) return parts[0] * 60 + parts[1];
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  return 0;
}

function extractYouTubeId(url) {
  if (!url) return null;
  const patterns = [
    /(?:youtube\.com\/watch\?v=|youtu\.be\/|youtube\.com\/embed\/|youtube\.com\/shorts\/)([a-zA-Z0-9_-]{11})/,
    /^([a-zA-Z0-9_-]{11})$/,
  ];
  for (const p of patterns) {
    const m = url.match(p);
    if (m) return m[1];
  }
  return null;
}

// ===== Spotify helpers =====
function extractSpotifyId(url, kind) {
  // kind: 'track' | 'playlist'
  if (!url) return null;
  const re = new RegExp(`(?:spotify[:/])${kind}[/:]([a-zA-Z0-9]{22})`, 'i');
  const m = url.match(re);
  if (m) return m[1];
  // Bare ID
  if (/^[a-zA-Z0-9]{22}$/.test(url.trim())) return url.trim();
  return null;
}

function spotifyAuthUrl(clientId, redirectUri) {
  const scopes = [
    'streaming',
    'user-read-email',
    'user-read-private',
    'user-modify-playback-state',
    'user-read-playback-state',
    'playlist-read-private',
    'playlist-read-collaborative',
  ].join(' ');
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: 'token',
    redirect_uri: redirectUri,
    scope: scopes,
    show_dialog: 'false',
  });
  return `https://accounts.spotify.com/authorize?${params.toString()}`;
}

function parseSpotifyHashToken() {
  if (!window.location.hash) return null;
  const params = new URLSearchParams(window.location.hash.slice(1));
  const token = params.get('access_token');
  if (!token) return null;
  const expiresIn = parseInt(params.get('expires_in') || '3600', 10);
  // Clean the URL
  window.history.replaceState({}, document.title, window.location.pathname + window.location.search);
  return { token, expiresAt: Date.now() + expiresIn * 1000 };
}

let spotifySdkPromise = null;
function loadSpotifySDK() {
  if (spotifySdkPromise) return spotifySdkPromise;
  spotifySdkPromise = new Promise((resolve) => {
    if (window.Spotify) return resolve(window.Spotify);
    window.onSpotifyWebPlaybackSDKReady = () => resolve(window.Spotify);
    const tag = document.createElement('script');
    tag.src = 'https://sdk.scdn.co/spotify-player.js';
    document.head.appendChild(tag);
  });
  return spotifySdkPromise;
}

// ===== YouTube IFrame API =====
let ytApiPromise = null;
function loadYouTubeAPI() {
  if (ytApiPromise) return ytApiPromise;
  ytApiPromise = new Promise((resolve) => {
    if (window.YT && window.YT.Player) return resolve(window.YT);
    const tag = document.createElement('script');
    tag.src = 'https://www.youtube.com/iframe_api';
    window.onYouTubeIframeAPIReady = () => resolve(window.YT);
    document.head.appendChild(tag);
  });
  return ytApiPromise;
}

export default function App() {
  const [questions, setQuestions] = useState(DEFAULT_QUESTIONS);
  const [readingTime, setReadingTime] = useState(300);
  const [examTitle, setExamTitle] = useState('Trinity School — Music Junior Form — Summer 2026');
  const [script, setScript] = useState(DEFAULT_SCRIPT);
  const [showScript, setShowScript] = useState(false);

  const [ttsProvider, setTtsProvider] = useState(() => localStorage.getItem('aural_tts_provider') || 'browser');
  const [elevenKey, setElevenKey] = useState(() => localStorage.getItem('aural_eleven_key') || '');
  const [elevenVoiceId, setElevenVoiceId] = useState(() => localStorage.getItem('aural_eleven_voice') || 'EXAVITQu4vr4xnSDxMaL');
  const [openaiKey, setOpenaiKey] = useState(() => localStorage.getItem('aural_openai_key') || '');
  const [openaiVoice, setOpenaiVoice] = useState(() => localStorage.getItem('aural_openai_voice') || 'alloy');
  const [browserVoices, setBrowserVoices] = useState([]);
  const [browserVoiceName, setBrowserVoiceName] = useState('');
  const [speechRate, setSpeechRate] = useState(0.95);
  const [speechPitch, setSpeechPitch] = useState(1);
  const [ttsTestStatus, setTtsTestStatus] = useState(null);

  const [previewingId, setPreviewingId] = useState(null);
  const [isCompiling, setIsCompiling] = useState(false);
  const [compileProgress, setCompileProgress] = useState(0);
  const [compileStatus, setCompileStatus] = useState('');
  const [finalAudioUrl, setFinalAudioUrl] = useState(null);
  const [finalAudioDuration, setFinalAudioDuration] = useState(0);
  const [showSettings, setShowSettings] = useState(false);
  const [shortReadingForTesting, setShortReadingForTesting] = useState(false);
  const [livePlaying, setLivePlaying] = useState(false);

  // ===== Spotify state =====
  const [spotifyClientId, setSpotifyClientId] = useState(() => localStorage.getItem('aural_spotify_client_id') || '');
  const [spotifyToken, setSpotifyToken] = useState(() => {
    const raw = localStorage.getItem('aural_spotify_token');
    if (!raw) return null;
    try {
      const parsed = JSON.parse(raw);
      if (parsed.expiresAt && parsed.expiresAt > Date.now() + 60_000) return parsed;
    } catch (e) {}
    return null;
  });
  const [spotifyUser, setSpotifyUser] = useState(null);
  const [spotifyDeviceId, setSpotifyDeviceId] = useState(null);
  const [spotifySdkReady, setSpotifySdkReady] = useState(false);
  const [spotifyImportedTracks, setSpotifyImportedTracks] = useState([]); // staging area for playlist tracks
  const [spotifyPlaylistName, setSpotifyPlaylistName] = useState('');
  const [spotifyPlaylistUrl, setSpotifyPlaylistUrl] = useState('');
  const [spotifyLoading, setSpotifyLoading] = useState(false);

  const audioContextRef = useRef(null);
  const previewStopRef = useRef(null);
  const ytPlayerRef = useRef(null);
  const ytContainerRef = useRef(null);
  const ttsCacheRef = useRef(new Map());
  const liveStopRef = useRef({ stopped: false });
  const spotifyPlayerRef = useRef(null);
  const spotifyPreviewBufferCache = useRef(new Map()); // url -> AudioBuffer

  // Capture Spotify OAuth callback on mount
  useEffect(() => {
    const tok = parseSpotifyHashToken();
    if (tok) {
      setSpotifyToken(tok);
      localStorage.setItem('aural_spotify_token', JSON.stringify(tok));
    }
  }, []);

  // Persist client ID
  useEffect(() => {
    if (spotifyClientId) localStorage.setItem('aural_spotify_client_id', spotifyClientId);
  }, [spotifyClientId]);

  // Fetch Spotify user profile when token is available
  useEffect(() => {
    if (!spotifyToken) { setSpotifyUser(null); return; }
    if (spotifyToken.expiresAt <= Date.now()) { setSpotifyToken(null); localStorage.removeItem('aural_spotify_token'); return; }
    fetch('https://api.spotify.com/v1/me', { headers: { Authorization: `Bearer ${spotifyToken.token}` } })
      .then(r => r.ok ? r.json() : null)
      .then(setSpotifyUser)
      .catch(() => setSpotifyUser(null));
  }, [spotifyToken]);

  // Initialise Spotify Web Playback SDK when token available
  useEffect(() => {
    if (!spotifyToken) return;
    let cancelled = false;
    (async () => {
      const Spotify = await loadSpotifySDK();
      if (cancelled) return;
      if (spotifyPlayerRef.current) return;
      const player = new Spotify.Player({
        name: 'Aural Composer',
        getOAuthToken: cb => cb(spotifyToken.token),
        volume: 1.0,
      });
      player.addListener('ready', ({ device_id }) => {
        setSpotifyDeviceId(device_id);
        setSpotifySdkReady(true);
      });
      player.addListener('not_ready', () => setSpotifySdkReady(false));
      player.addListener('authentication_error', ({ message }) => {
        console.warn('Spotify auth error:', message);
        setSpotifyToken(null);
        localStorage.removeItem('aural_spotify_token');
      });
      player.addListener('account_error', ({ message }) => {
        alert(`Spotify account error: ${message}\n\nFull-track playback requires a Premium account.`);
      });
      await player.connect();
      spotifyPlayerRef.current = player;
    })();
    return () => { cancelled = true; };
  }, [spotifyToken]);

  useEffect(() => { localStorage.setItem('aural_tts_provider', ttsProvider); }, [ttsProvider]);
  useEffect(() => { localStorage.setItem('aural_eleven_key', elevenKey); }, [elevenKey]);
  useEffect(() => { localStorage.setItem('aural_eleven_voice', elevenVoiceId); }, [elevenVoiceId]);
  useEffect(() => { localStorage.setItem('aural_openai_key', openaiKey); }, [openaiKey]);
  useEffect(() => { localStorage.setItem('aural_openai_voice', openaiVoice); }, [openaiVoice]);

  useEffect(() => {
    function loadVoices() {
      const v = window.speechSynthesis.getVoices();
      const sorted = v.slice().sort((a, b) => {
        const aEn = a.lang.startsWith('en');
        const bEn = b.lang.startsWith('en');
        if (aEn !== bEn) return bEn - aEn;
        const aGB = a.lang === 'en-GB';
        const bGB = b.lang === 'en-GB';
        if (aGB !== bGB) return bGB - aGB;
        return a.name.localeCompare(b.name);
      });
      setBrowserVoices(sorted);
      if (sorted.length && !browserVoiceName) {
        const preferred = sorted.find(x => /sonia|libby|amelia|female/i.test(x.name) && x.lang === 'en-GB') || sorted.find(x => x.lang === 'en-GB') || sorted[0];
        setBrowserVoiceName(preferred.name);
      }
    }
    loadVoices();
    window.speechSynthesis.onvoiceschanged = loadVoices;
    return () => { window.speechSynthesis.onvoiceschanged = null; };
  }, []);

  useEffect(() => { loadYouTubeAPI(); }, []);

  const getAudioContext = () => {
    if (!audioContextRef.current) {
      audioContextRef.current = new (window.AudioContext || window.webkitAudioContext)();
    }
    return audioContextRef.current;
  };

  const setSource = (questionId, source) => {
    setQuestions(prev => prev.map(q => q.id === questionId ? { ...q, source } : q));
  };

  const handleFileUpload = async (questionId, file) => {
    if (!file) return;
    try {
      const ctx = getAudioContext();
      const arrayBuffer = await file.arrayBuffer();
      const audioBuffer = await ctx.decodeAudioData(arrayBuffer.slice(0));
      setSource(questionId, { kind: 'file', name: file.name, buffer: audioBuffer, duration: audioBuffer.duration });
    } catch (err) {
      alert(`Could not decode audio file: ${err.message}`);
    }
  };

  const handleYouTubeSet = (questionId, { url, startStr, endStr }) => {
    const videoId = extractYouTubeId(url);
    if (!videoId) { alert('Could not parse YouTube URL.'); return; }
    const start = parseTimestamp(startStr);
    const end = parseTimestamp(endStr);
    if (end <= start) { alert('End timestamp must be later than start timestamp.'); return; }
    setSource(questionId, { kind: 'youtube', videoId, url, start, end, duration: end - start, startStr, endStr });
  };

  // ===== Spotify handlers =====
  const spotifyConnect = () => {
    if (!spotifyClientId) {
      alert('Please enter your Spotify Client ID first (Voice & API → Spotify).');
      return;
    }
    const redirectUri = window.location.origin + window.location.pathname;
    const authUrl = spotifyAuthUrl(spotifyClientId, redirectUri);
    window.location.href = authUrl;
  };

  const spotifyDisconnect = () => {
    setSpotifyToken(null);
    setSpotifyUser(null);
    setSpotifyDeviceId(null);
    setSpotifySdkReady(false);
    localStorage.removeItem('aural_spotify_token');
    if (spotifyPlayerRef.current) {
      try { spotifyPlayerRef.current.disconnect(); } catch (e) {}
      spotifyPlayerRef.current = null;
    }
  };

  const spotifyFetch = async (url) => {
    if (!spotifyToken) throw new Error('Not connected to Spotify');
    const res = await fetch(url, { headers: { Authorization: `Bearer ${spotifyToken.token}` } });
    if (res.status === 401) {
      spotifyDisconnect();
      throw new Error('Spotify session expired. Please reconnect.');
    }
    if (!res.ok) throw new Error(`Spotify API error: ${res.status}`);
    return res.json();
  };

  const importSpotifyPlaylist = async () => {
    const playlistId = extractSpotifyId(spotifyPlaylistUrl, 'playlist');
    if (!playlistId) { alert('Could not parse Spotify playlist URL/ID.'); return; }
    setSpotifyLoading(true);
    try {
      const meta = await spotifyFetch(`https://api.spotify.com/v1/playlists/${playlistId}?fields=name,description,tracks.total`);
      setSpotifyPlaylistName(meta.name);

      let allTracks = [];
      let next = `https://api.spotify.com/v1/playlists/${playlistId}/tracks?fields=items(track(id,name,artists(name),duration_ms,preview_url,uri,external_urls)),next&limit=100`;
      while (next) {
        const page = await spotifyFetch(next);
        for (const item of page.items) {
          if (item.track && item.track.id) {
            allTracks.push({
              id: item.track.id,
              name: item.track.name,
              artists: item.track.artists.map(a => a.name).join(', '),
              durationMs: item.track.duration_ms,
              previewUrl: item.track.preview_url,
              uri: item.track.uri,
              externalUrl: item.track.external_urls?.spotify,
            });
          }
        }
        next = page.next;
      }
      setSpotifyImportedTracks(allTracks);
    } catch (err) {
      alert(`Failed to import playlist: ${err.message}`);
    } finally {
      setSpotifyLoading(false);
    }
  };

  const importSpotifyTrack = async () => {
    const trackId = extractSpotifyId(spotifyPlaylistUrl, 'track');
    if (!trackId) { alert('Could not parse Spotify track URL/ID.'); return; }
    setSpotifyLoading(true);
    try {
      const t = await spotifyFetch(`https://api.spotify.com/v1/tracks/${trackId}`);
      setSpotifyImportedTracks([{
        id: t.id,
        name: t.name,
        artists: t.artists.map(a => a.name).join(', '),
        durationMs: t.duration_ms,
        previewUrl: t.preview_url,
        uri: t.uri,
        externalUrl: t.external_urls?.spotify,
      }]);
      setSpotifyPlaylistName('Single track');
    } catch (err) {
      alert(`Failed to import track: ${err.message}`);
    } finally {
      setSpotifyLoading(false);
    }
  };

  const assignSpotifyTrackToQuestion = (questionId, track, startStr = '0:00', endStr = null) => {
    const start = parseTimestamp(startStr);
    const trackDurSec = track.durationMs / 1000;
    const end = endStr ? parseTimestamp(endStr) : trackDurSec;
    if (end <= start) { alert('End must be later than start.'); return; }
    if (end > trackDurSec + 0.5) { alert(`Track is only ${formatTime(trackDurSec)} long. End time exceeds track length.`); return; }

    setSource(questionId, {
      kind: 'spotify',
      trackId: track.id,
      uri: track.uri,
      name: track.name,
      artists: track.artists,
      durationMs: track.durationMs,
      previewUrl: track.previewUrl,
      externalUrl: track.externalUrl,
      start,
      end,
      duration: end - start,
      startStr,
      endStr: endStr || formatTime(end),
    });
  };

  // Fetch a single Spotify track by URL/URI and assign to a question (used by per-extract Spotify tab)
  const handleSpotifyTrackAdd = async (questionId, url, startStr, endStr) => {
    const trackId = extractSpotifyId(url, 'track');
    if (!trackId) { alert('Could not parse Spotify track URL.'); return; }
    try {
      const t = await spotifyFetch(`https://api.spotify.com/v1/tracks/${trackId}`);
      const track = {
        id: t.id, name: t.name,
        artists: t.artists.map(a => a.name).join(', '),
        durationMs: t.duration_ms, previewUrl: t.preview_url, uri: t.uri,
        externalUrl: t.external_urls?.spotify,
      };
      assignSpotifyTrackToQuestion(questionId, track, startStr || '0:00', endStr || null);
    } catch (err) {
      alert(`Could not load track: ${err.message}`);
    }
  };

  // Play a Spotify segment via the Web Playback SDK
  const playSpotifySegment = async (uri, startSec, endSec) => {
    if (!spotifyDeviceId || !spotifyToken) {
      alert('Spotify player not ready. Connect to Spotify in Voice & API settings.');
      return;
    }
    // Transfer playback to our device & start at offset
    const startMs = Math.floor(startSec * 1000);
    const playRes = await fetch(`https://api.spotify.com/v1/me/player/play?device_id=${spotifyDeviceId}`, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${spotifyToken.token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ uris: [uri], position_ms: startMs }),
    });
    if (!playRes.ok && playRes.status !== 204) {
      const errText = await playRes.text();
      throw new Error(`Spotify play failed: ${playRes.status} ${errText.slice(0, 150)}`);
    }
    // Wait until end time
    const durationMs = (endSec - startSec) * 1000;
    await new Promise(r => setTimeout(r, durationMs));
    // Pause
    try {
      await fetch(`https://api.spotify.com/v1/me/player/pause?device_id=${spotifyDeviceId}`, {
        method: 'PUT',
        headers: { Authorization: `Bearer ${spotifyToken.token}` },
      });
    } catch (e) {}
  };

  // Load and decode a Spotify preview URL into an AudioBuffer (used for WAV export when clip fits in 30s preview)
  const loadSpotifyPreviewBuffer = async (previewUrl) => {
    if (!previewUrl) return null;
    if (spotifyPreviewBufferCache.current.has(previewUrl)) return spotifyPreviewBufferCache.current.get(previewUrl);
    try {
      const res = await fetch(previewUrl);
      if (!res.ok) return null;
      const arrayBuffer = await res.arrayBuffer();
      const buf = await getAudioContext().decodeAudioData(arrayBuffer);
      spotifyPreviewBufferCache.current.set(previewUrl, buf);
      return buf;
    } catch (e) {
      console.warn('Preview load failed:', e);
      return null;
    }
  };

  const clearSource = (questionId) => setSource(questionId, null);
  const updateQuestion = (id, field, value) => {
    setQuestions(prev => prev.map(q => q.id === id ? { ...q, [field]: value } : q));
  };

  // ===== Question CRUD =====
  const addQuestion = (afterIndex = null) => {
    setQuestions(prev => {
      const maxId = prev.reduce((m, q) => Math.max(m, q.id), 0);
      const newQ = {
        id: maxId + 1,
        label: `Extract ${prev.length + 1}`,
        plays: 3,
        gapBetweenPlays: 30,
        gapAfter: 45,
        marks: null,
        intro: buildIntroForExtract(`Extract ${prev.length + 1}`, 3),
        source: null,
      };
      if (afterIndex === null || afterIndex >= prev.length - 1) return [...prev, newQ];
      const copy = [...prev];
      copy.splice(afterIndex + 1, 0, newQ);
      return copy;
    });
  };

  const deleteQuestion = (id) => {
    if (!confirm('Delete this extract?')) return;
    setQuestions(prev => prev.filter(q => q.id !== id));
  };

  const moveQuestion = (id, direction) => {
    setQuestions(prev => {
      const idx = prev.findIndex(q => q.id === id);
      if (idx === -1) return prev;
      const newIdx = direction === 'up' ? idx - 1 : idx + 1;
      if (newIdx < 0 || newIdx >= prev.length) return prev;
      const copy = [...prev];
      [copy[idx], copy[newIdx]] = [copy[newIdx], copy[idx]];
      return copy;
    });
  };

  const moveQuestionTo = (fromIndex, toIndex) => {
    setQuestions(prev => {
      if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0 || fromIndex >= prev.length || toIndex >= prev.length) return prev;
      const copy = [...prev];
      const [item] = copy.splice(fromIndex, 1);
      copy.splice(toIndex, 0, item);
      return copy;
    });
  };

  // ===== PDF parsing =====
  const [pdfParsing, setPdfParsing] = useState(false);
  const [pdfDetectionInfo, setPdfDetectionInfo] = useState(null); // { title, count, marksFound }

  const handleExamPdfDrop = async (file) => {
    if (!file) return;
    if (!file.name.toLowerCase().endsWith('.pdf')) {
      alert('Please drop a PDF file.');
      return;
    }
    setPdfParsing(true);
    setPdfDetectionInfo(null);
    try {
      const result = await parseExamPdf(file);
      if (result.extracts.length === 0) {
        alert('No "Extract N" or "Question N" headings found in this PDF. You can still add extracts manually.');
        setPdfParsing(false);
        return;
      }

      // Confirm before overwriting if user already has data
      const anyFilled = questions.some(q => q.source);
      if (anyFilled) {
        if (!confirm(`Detected ${result.extracts.length} extracts in the PDF.\n\nThis will replace your current extract list (uploaded audio will be cleared). Continue?`)) {
          setPdfParsing(false);
          return;
        }
      }

      const newQuestions = result.extracts.map((ex, i) => {
        const isLast = i === result.extracts.length - 1;
        return {
          id: i + 1,
          label: ex.label,
          plays: ex.plays,
          gapBetweenPlays: ex.plays === 2 ? 20 : 30,
          gapAfter: isLast ? 30 : 60,
          marks: ex.marks,
          intro: buildIntroForExtract(ex.label, ex.plays, isLast),
          source: null,
        };
      });
      setQuestions(newQuestions);
      if (result.title) setExamTitle(result.title);
      setPdfDetectionInfo({
        count: result.extracts.length,
        marksFound: result.marksTableFound,
        title: result.title,
      });
    } catch (err) {
      console.error(err);
      alert(`Could not parse PDF: ${err.message}`);
    } finally {
      setPdfParsing(false);
    }
  };

  // ===== Save / load exam config =====
  const saveExamConfig = () => {
    const config = {
      version: 1,
      savedAt: new Date().toISOString(),
      examTitle,
      readingTime,
      script,
      questions: questions.map(q => ({
        id: q.id,
        label: q.label,
        plays: q.plays,
        gapBetweenPlays: q.gapBetweenPlays,
        gapAfter: q.gapAfter,
        marks: q.marks,
        intro: q.intro,
        // Sources are NOT saved (audio buffers are too big; URLs/IDs are saved for non-file sources)
        source: q.source && q.source.kind !== 'file' ? q.source : null,
      })),
    };
    const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${examTitle.replace(/[^a-z0-9]+/gi, '_')}.aural.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const loadExamConfig = async (file) => {
    if (!file) return;
    try {
      const text = await file.text();
      const config = JSON.parse(text);
      if (!config.questions || !Array.isArray(config.questions)) throw new Error('Invalid config file');
      if (!confirm('Loading a saved config will replace your current setup (uploaded audio will be cleared). Continue?')) return;
      setExamTitle(config.examTitle || '');
      setReadingTime(config.readingTime || 300);
      if (config.script) setScript(config.script);
      setQuestions(config.questions.map(q => ({
        ...q,
        // Restored sources retain their kind/url/timestamps but no decoded buffer
        source: q.source && q.source.kind !== 'file' ? q.source : null,
      })));
    } catch (err) {
      alert(`Could not load config: ${err.message}`);
    }
  };

  const estimateSpeechDuration = (text) => {
    const words = text.trim().split(/\s+/).length;
    return (words / 150) * 60 + 0.8;
  };

  const cacheKey = (text) => `${ttsProvider}|${ttsProvider === 'eleven' ? elevenVoiceId : ttsProvider === 'openai' ? openaiVoice : browserVoiceName}|${speechRate}|${text}`;

  const renderTTSBuffer = async (text) => {
    const key = cacheKey(text);
    if (ttsCacheRef.current.has(key)) return ttsCacheRef.current.get(key);
    let buffer = null;
    if (ttsProvider === 'eleven' && elevenKey) buffer = await renderElevenLabs(text);
    else if (ttsProvider === 'openai' && openaiKey) buffer = await renderOpenAI(text);
    if (buffer) ttsCacheRef.current.set(key, buffer);
    return buffer;
  };

  const renderElevenLabs = async (text) => {
    const res = await fetch(`https://api.elevenlabs.io/v1/text-to-speech/${elevenVoiceId}`, {
      method: 'POST',
      headers: { 'xi-api-key': elevenKey, 'Content-Type': 'application/json', 'Accept': 'audio/mpeg' },
      body: JSON.stringify({
        text,
        model_id: 'eleven_multilingual_v2',
        voice_settings: { stability: 0.55, similarity_boost: 0.75, style: 0.0 },
      }),
    });
    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`ElevenLabs error (${res.status}): ${errText.slice(0, 200)}`);
    }
    const arrayBuffer = await res.arrayBuffer();
    return await getAudioContext().decodeAudioData(arrayBuffer);
  };

  const renderOpenAI = async (text) => {
    const res = await fetch('https://api.openai.com/v1/audio/speech', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${openaiKey}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ model: 'tts-1-hd', voice: openaiVoice, input: text, speed: speechRate, response_format: 'mp3' }),
    });
    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`OpenAI error (${res.status}): ${errText.slice(0, 200)}`);
    }
    const arrayBuffer = await res.arrayBuffer();
    return await getAudioContext().decodeAudioData(arrayBuffer);
  };

  const speakLive = async (text) => {
    if (ttsProvider === 'browser' || (ttsProvider === 'eleven' && !elevenKey) || (ttsProvider === 'openai' && !openaiKey)) {
      return new Promise((resolve) => {
        const utter = new SpeechSynthesisUtterance(text);
        const voice = browserVoices.find(v => v.name === browserVoiceName);
        if (voice) utter.voice = voice;
        utter.rate = speechRate;
        utter.pitch = speechPitch;
        utter.onend = resolve;
        utter.onerror = resolve;
        window.speechSynthesis.speak(utter);
      });
    }
    try {
      const buf = await renderTTSBuffer(text);
      if (!buf) throw new Error('No buffer');
      const ctx = getAudioContext();
      if (ctx.state === 'suspended') await ctx.resume();
      const src = ctx.createBufferSource();
      src.buffer = buf;
      src.connect(ctx.destination);
      src.start();
      return new Promise((resolve) => { src.onended = resolve; });
    } catch (err) {
      console.warn('Premium TTS failed, falling back to browser:', err);
      return new Promise((resolve) => {
        const utter = new SpeechSynthesisUtterance(text);
        const voice = browserVoices.find(v => v.name === browserVoiceName);
        if (voice) utter.voice = voice;
        utter.rate = speechRate;
        utter.onend = resolve;
        utter.onerror = resolve;
        window.speechSynthesis.speak(utter);
      });
    }
  };

  const testVoice = async () => {
    setTtsTestStatus('loading');
    try {
      await speakLive('This is a test of the examiner voice. You will hear this extract three times.');
      setTtsTestStatus('ok');
      setTimeout(() => setTtsTestStatus(null), 2000);
    } catch (err) {
      setTtsTestStatus('error');
      alert(`Voice test failed: ${err.message}`);
    }
  };

  const buildTimeline = useCallback(() => {
    const timeline = [];
    let cursor = 0;
    const effectiveReading = shortReadingForTesting ? 10 : readingTime;
    const addTTS = (text) => { const d = estimateSpeechDuration(text); timeline.push({ type: 'tts', text, start: cursor, duration: d }); cursor += d; };
    const addSilence = (d, label) => { timeline.push({ type: 'silence', start: cursor, duration: d, label }); cursor += d; };

    addTTS(script.opening);
    addSilence(effectiveReading, 'Reading time');
    addTTS(script.postReading);
    addSilence(2);

    questions.forEach((q, qi) => {
      addTTS(q.intro);
      addSilence(1.5);

      if (q.source) {
        for (let i = 0; i < q.plays; i++) {
          if (q.source.kind === 'file') {
            timeline.push({ type: 'audio', buffer: q.source.buffer, start: cursor, duration: q.source.duration, questionId: q.id, playNumber: i + 1 });
            cursor += q.source.duration;
          } else if (q.source.kind === 'youtube') {
            timeline.push({ type: 'youtube', videoId: q.source.videoId, startSec: q.source.start, endSec: q.source.end, start: cursor, duration: q.source.duration, questionId: q.id, playNumber: i + 1 });
            cursor += q.source.duration;
          } else if (q.source.kind === 'spotify') {
            timeline.push({
              type: 'spotify',
              uri: q.source.uri,
              previewUrl: q.source.previewUrl,
              startSec: q.source.start,
              endSec: q.source.end,
              start: cursor,
              duration: q.source.duration,
              questionId: q.id,
              playNumber: i + 1,
            });
            cursor += q.source.duration;
          }
          if (i < q.plays - 1) {
            const upcomingPlayNumber = i + 2; // next time the extract plays
            const isFinal = upcomingPlayNumber === q.plays;
            const announce = renderBetweenPlays(script.betweenPlays, upcomingPlayNumber, isFinal);
            const announceDur = estimateSpeechDuration(announce);
            timeline.push({ type: 'tts', text: announce, start: cursor, duration: announceDur });
            cursor += announceDur;
            const remaining = Math.max(0, q.gapBetweenPlays - announceDur);
            addSilence(remaining, 'Thinking time');
          }
        }
      } else {
        addSilence(10, '[No audio set]');
      }
      if (qi < questions.length - 1) addSilence(q.gapAfter, 'Answer time');
    });

    addTTS(script.ending);
    return { timeline, totalDuration: cursor };
  }, [questions, readingTime, shortReadingForTesting, script]);

  const previewQuestion = async (q) => {
    if (previewingId === q.id) { stopAll(); return; }
    setPreviewingId(q.id);
    try {
      // Find the question's index to know if it's the last
      const qIdx = questions.findIndex(x => x.id === q.id);
      const isLast = qIdx === questions.length - 1;

      // Intro
      await speakLive(q.intro);

      // All plays + between announcements
      for (let i = 0; i < q.plays; i++) {
        if (q.source?.kind === 'file') {
          await new Promise((resolve) => {
            const ctx = getAudioContext();
            const src = ctx.createBufferSource();
            src.buffer = q.source.buffer;
            src.connect(ctx.destination);
            src.start();
            previewStopRef.current = { stop: () => { try { src.stop(); } catch(e){} } };
            src.onended = resolve;
          });
        } else if (q.source?.kind === 'youtube') {
          await playYouTubeSegment(q.source.videoId, q.source.start, q.source.end);
        } else if (q.source?.kind === 'spotify') {
          await playSpotifySegment(q.source.uri, q.source.start, q.source.end);
        }
        // Between-plays announcement (not after last play)
        if (i < q.plays - 1) {
          const upcoming = i + 2;
          const announce = renderBetweenPlays(script.betweenPlays, upcoming, upcoming === q.plays);
          await speakLive(announce);
          // Short pause to simulate thinking time, capped at 5s for preview brevity
          const pauseMs = Math.min(q.gapBetweenPlays * 1000, 5000);
          await new Promise(r => setTimeout(r, pauseMs));
        }
      }
    } finally {
      setPreviewingId(null);
    }
  };

  const stopAll = () => {
    window.speechSynthesis.cancel();
    if (previewStopRef.current) { previewStopRef.current.stop?.(); previewStopRef.current = null; }
    if (ytPlayerRef.current) { try { ytPlayerRef.current.stopVideo(); } catch (e) {} }
    if (spotifyPlayerRef.current) { try { spotifyPlayerRef.current.pause(); } catch (e) {} }
    liveStopRef.current.stopped = true;
    setPreviewingId(null);
    setLivePlaying(false);
  };

  const ensureYouTubePlayer = async () => {
    await loadYouTubeAPI();
    if (ytPlayerRef.current) return ytPlayerRef.current;
    return new Promise((resolve) => {
      ytPlayerRef.current = new window.YT.Player(ytContainerRef.current, {
        height: '120', width: '200',
        playerVars: { controls: 0, disablekb: 1, modestbranding: 1, rel: 0 },
        events: { onReady: () => resolve(ytPlayerRef.current) },
      });
    });
  };

  const playYouTubeSegment = async (videoId, startSec, endSec) => {
    const player = await ensureYouTubePlayer();
    return new Promise((resolve) => {
      let resolved = false;
      const done = () => { if (!resolved) { resolved = true; clearInterval(checker); resolve(); } };
      player.loadVideoById({ videoId, startSeconds: startSec, endSeconds: endSec });
      player.unMute();
      player.setVolume(100);
      const checker = setInterval(() => {
        if (resolved) return;
        try {
          const t = player.getCurrentTime();
          const state = player.getPlayerState();
          if (state === 0 || t >= endSec - 0.1) {
            player.pauseVideo();
            done();
          }
        } catch (e) { }
      }, 150);
      setTimeout(done, (endSec - startSec + 5) * 1000);
    });
  };

  const playLiveFull = async () => {
    if (livePlaying) {
      liveStopRef.current.stopped = true;
      stopAll();
      return;
    }
    const { timeline } = buildTimeline();
    liveStopRef.current = { stopped: false };
    setLivePlaying(true);

    const ctx = getAudioContext();
    if (ctx.state === 'suspended') await ctx.resume();
    const startedAt = performance.now() / 1000;

    for (const item of timeline) {
      if (liveStopRef.current.stopped) break;
      const targetTime = startedAt + item.start;
      const now = performance.now() / 1000;
      const wait = (targetTime - now) * 1000;
      if (wait > 0) await new Promise(r => setTimeout(r, wait));
      if (liveStopRef.current.stopped) break;

      if (item.type === 'tts') {
        await speakLive(item.text);
      } else if (item.type === 'audio') {
        await new Promise((resolve) => {
          const src = ctx.createBufferSource();
          src.buffer = item.buffer;
          src.connect(ctx.destination);
          src.start();
          src.onended = resolve;
        });
      } else if (item.type === 'youtube') {
        await playYouTubeSegment(item.videoId, item.startSec, item.endSec);
      } else if (item.type === 'spotify') {
        await playSpotifySegment(item.uri, item.startSec, item.endSec);
      }
    }
    setLivePlaying(false);
  };

  const compileAudio = async () => {
    const filled = questions.filter(q => q.source);
    if (filled.length === 0) { alert('Add audio to at least one extract before compiling.'); return; }

    const hasYouTube = filled.some(q => q.source.kind === 'youtube');
    const hasSpotify = filled.some(q => q.source.kind === 'spotify');
    const usingPremiumTTS = (ttsProvider === 'eleven' && elevenKey) || (ttsProvider === 'openai' && openaiKey);

    if (hasYouTube || hasSpotify) {
      const lines = [];
      if (hasYouTube) lines.push('• YouTube clips cannot be embedded in the WAV (DRM). They will be silence in the file.');
      if (hasSpotify) lines.push('• Spotify tracks cannot be embedded in the WAV (DRM). If a 30-second preview clip is available AND your start/end timestamps fall inside the first 30 seconds, that portion will be baked in. Otherwise it will be silence.');
      lines.push('');
      lines.push('All sources play correctly during "Preview full exam (live)". Continue compiling?');
      const ok = confirm(lines.join('\n'));
      if (!ok) return;
    }

    setIsCompiling(true);
    setCompileProgress(0);
    setCompileStatus('Building timeline...');

    try {
      const { timeline, totalDuration } = buildTimeline();
      const sampleRate = 44100;
      const numChannels = 2;
      const totalSamples = Math.ceil(totalDuration * sampleRate);
      const offlineCtx = new OfflineAudioContext(numChannels, totalSamples, sampleRate);

      const ttsItems = timeline.filter(t => t.type === 'tts');
      const ttsBuffers = new Map();

      if (usingPremiumTTS) {
        for (let i = 0; i < ttsItems.length; i++) {
          setCompileStatus(`Generating announcement ${i + 1} of ${ttsItems.length}...`);
          setCompileProgress((i / ttsItems.length) * 70);
          try {
            const buf = await renderTTSBuffer(ttsItems[i].text);
            ttsBuffers.set(i, buf);
          } catch (err) {
            console.warn('TTS failed:', err);
            ttsBuffers.set(i, null);
          }
        }
      } else {
        setCompileStatus('Browser TTS cannot be recorded — inserting timing markers...');
        for (let i = 0; i < ttsItems.length; i++) {
          ttsBuffers.set(i, await makeMarkerTone(ttsItems[i].duration));
        }
      }

      setCompileStatus('Stitching audio...');
      setCompileProgress(80);

      // Pre-load any spotify preview buffers that fit within their clip range
      const spotifyItems = timeline.filter(t => t.type === 'spotify' && t.previewUrl && t.endSec <= 30);
      const spotifyPreviewBuffers = new Map();
      for (const item of spotifyItems) {
        try {
          const buf = await loadSpotifyPreviewBuffer(item.previewUrl);
          if (buf) spotifyPreviewBuffers.set(item.previewUrl, buf);
        } catch (e) {}
      }

      let ttsIndex = 0;
      timeline.forEach((item) => {
        if (item.type === 'audio') {
          const src = offlineCtx.createBufferSource();
          src.buffer = item.buffer;
          src.connect(offlineCtx.destination);
          src.start(item.start);
        } else if (item.type === 'tts') {
          const buf = ttsBuffers.get(ttsIndex);
          ttsIndex++;
          if (buf) {
            const src = offlineCtx.createBufferSource();
            src.buffer = buf;
            src.connect(offlineCtx.destination);
            src.start(item.start);
          }
        } else if (item.type === 'spotify' && item.previewUrl && item.endSec <= 30 && spotifyPreviewBuffers.has(item.previewUrl)) {
          const buf = spotifyPreviewBuffers.get(item.previewUrl);
          const src = offlineCtx.createBufferSource();
          src.buffer = buf;
          src.connect(offlineCtx.destination);
          // Offset within preview clip = item.startSec
          src.start(item.start, item.startSec, item.duration);
        }
      });

      setCompileStatus('Rendering...');
      setCompileProgress(90);
      const rendered = await offlineCtx.startRendering();
      setCompileStatus('Encoding WAV...');
      setCompileProgress(97);
      const wavBlob = audioBufferToWav(rendered);
      const url = URL.createObjectURL(wavBlob);
      if (finalAudioUrl) URL.revokeObjectURL(finalAudioUrl);
      setFinalAudioUrl(url);
      setFinalAudioDuration(rendered.duration);
      setCompileProgress(100);
      setCompileStatus('Done!');
      setTimeout(() => setIsCompiling(false), 600);
    } catch (err) {
      console.error(err);
      alert(`Compile failed: ${err.message}`);
      setIsCompiling(false);
    }
  };

  const makeMarkerTone = async (duration) => {
    const ctx = new OfflineAudioContext(2, Math.ceil(duration * 44100), 44100);
    const osc = ctx.createOscillator();
    const gain = ctx.createGain();
    osc.frequency.value = 880;
    gain.gain.setValueAtTime(0, 0);
    gain.gain.linearRampToValueAtTime(0.05, 0.02);
    gain.gain.linearRampToValueAtTime(0, 0.2);
    osc.connect(gain).connect(ctx.destination);
    osc.start(0); osc.stop(0.25);
    return await ctx.startRendering();
  };

  const downloadFinal = () => {
    if (!finalAudioUrl) return;
    const a = document.createElement('a');
    a.href = finalAudioUrl;
    a.download = `${examTitle.replace(/[^a-z0-9]+/gi, '_')}.wav`;
    a.click();
  };

  const { totalDuration } = buildTimeline();
  const filledCount = questions.filter(q => q.source).length;
  const youtubeCount = questions.filter(q => q.source?.kind === 'youtube').length;
  const spotifyCount = questions.filter(q => q.source?.kind === 'spotify').length;
  const totalMarks = questions.reduce((sum, q) => sum + (q.marks || 0), 0);

  return (
    <div className="min-h-screen" style={{
      background: 'linear-gradient(180deg, #f5f1e8 0%, #ede5d3 100%)',
      fontFamily: "'Source Serif Pro', 'Georgia', serif",
      color: '#2a2520',
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Source+Serif+Pro:ital,wght@0,400;0,600;0,700;1,400&family=JetBrains+Mono:wght@400;500&family=Cormorant+Garamond:ital,wght@0,400;0,600;0,700;1,400&display=swap');
        body { margin: 0; }
        .display-font { font-family: 'Cormorant Garamond', serif; }
        .mono-font { font-family: 'JetBrains Mono', monospace; }
        .ink-shadow { box-shadow: 0 1px 0 rgba(42,37,32,0.04), 0 4px 16px rgba(42,37,32,0.06), 0 0 0 1px rgba(42,37,32,0.08); }
        .accent { color: #8b2c1e; }
        .accent-bg { background: #8b2c1e; }
        .paper { background: #fdfbf5; }
        .hairline { border: 1px solid rgba(42,37,32,0.12); }
        input[type="number"], input[type="text"], input[type="password"], textarea, select {
          background: #fdfbf5; border: 1px solid rgba(42,37,32,0.15);
          padding: 6px 10px; font-family: inherit; color: #2a2520;
          border-radius: 2px; transition: border-color 0.15s;
        }
        input:focus, textarea:focus, select:focus { outline: none; border-color: #8b2c1e; }
        button { transition: all 0.15s; cursor: pointer; }
        button:disabled { opacity: 0.5; cursor: not-allowed; }
        .question-card { transition: all 0.2s; }
        .question-card:hover { transform: translateY(-1px); }
        .drop-zone.has-source { background: #f0e8d6; border-color: #8b2c1e; border-style: solid; }
        @keyframes shimmer { 0% { background-position: -200% 0; } 100% { background-position: 200% 0; } }
        .progress-bar {
          background: linear-gradient(90deg, #8b2c1e 0%, #c44530 50%, #8b2c1e 100%);
          background-size: 200% 100%; animation: shimmer 2s linear infinite;
        }
        .tab { padding: 6px 12px; border: 1px solid rgba(42,37,32,0.15); background: transparent; }
        .tab.active { background: #2a2520; color: #fdfbf5; border-color: #2a2520; }
        .yt-hidden { position: fixed; bottom: -200px; right: 10px; opacity: 0.01; pointer-events: none; }
        details > summary { list-style: none; cursor: pointer; }
        details > summary::-webkit-details-marker { display: none; }
      `}</style>

      <div className="yt-hidden"><div ref={ytContainerRef}></div></div>

      <header className="hairline" style={{ borderBottom: '2px solid #2a2520', background: '#fdfbf5' }}>
        <div className="max-w-6xl mx-auto px-8 py-6 flex items-center justify-between">
          <div>
            <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-1">Listening Examination Audio Compiler</div>
            <h1 className="display-font text-3xl font-bold leading-none">
              <span className="italic">Aural</span> · <span>Composer</span>
            </h1>
          </div>
          <button onClick={() => setShowSettings(!showSettings)}
            className="flex items-center gap-2 px-3 py-2 hairline"
            style={{ background: showSettings ? '#2a2520' : 'transparent', color: showSettings ? '#fdfbf5' : 'inherit' }}>
            <Settings size={14} />
            <span className="mono-font text-xs uppercase tracking-wider">Voice & API</span>
          </button>
        </div>
      </header>

      {showSettings && (
        <div className="paper hairline" style={{ borderBottom: '1px solid rgba(42,37,32,0.12)' }}>
          <div className="max-w-6xl mx-auto px-8 py-6">
            <div className="mb-4">
              <div className="mono-font text-xs uppercase tracking-wider opacity-60 mb-2">TTS Provider</div>
              <div className="flex gap-2 flex-wrap">
                <button className={`tab mono-font text-xs uppercase tracking-wider ${ttsProvider === 'browser' ? 'active' : ''}`} onClick={() => setTtsProvider('browser')}>Browser (free)</button>
                <button className={`tab mono-font text-xs uppercase tracking-wider ${ttsProvider === 'eleven' ? 'active' : ''}`} onClick={() => setTtsProvider('eleven')}>ElevenLabs</button>
                <button className={`tab mono-font text-xs uppercase tracking-wider ${ttsProvider === 'openai' ? 'active' : ''}`} onClick={() => setTtsProvider('openai')}>OpenAI</button>
              </div>
            </div>

            {ttsProvider === 'browser' && (
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 block mb-1">Browser voice</label>
                  <select value={browserVoiceName} onChange={e => setBrowserVoiceName(e.target.value)} className="w-full text-sm">
                    {browserVoices.map(v => <option key={v.name} value={v.name}>{v.name} ({v.lang})</option>)}
                  </select>
                </div>
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 block mb-1">Rate: <span className="accent font-semibold">{speechRate.toFixed(2)}×</span></label>
                  <input type="range" min="0.5" max="1.5" step="0.05" value={speechRate} onChange={e => setSpeechRate(parseFloat(e.target.value))} className="w-full" />
                </div>
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 block mb-1">Pitch: <span className="accent font-semibold">{speechPitch.toFixed(2)}</span></label>
                  <input type="range" min="0.5" max="1.5" step="0.05" value={speechPitch} onChange={e => setSpeechPitch(parseFloat(e.target.value))} className="w-full" />
                </div>
              </div>
            )}

            {ttsProvider === 'eleven' && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 flex items-center gap-1.5 mb-1"><Key size={11} /> API key</label>
                  <input type="password" value={elevenKey} onChange={e => setElevenKey(e.target.value)} placeholder="xi-..." className="w-full text-sm" />
                </div>
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 block mb-1">Voice</label>
                  <select value={elevenVoiceId} onChange={e => setElevenVoiceId(e.target.value)} className="w-full text-sm">
                    <option value="EXAVITQu4vr4xnSDxMaL">Sarah (en-US, calm)</option>
                    <option value="9BWtsMINqrJLrRacOk9x">Aria (en-US, warm)</option>
                    <option value="FGY2WhTYpPnrIDTdsKH5">Laura (en-US)</option>
                    <option value="cgSgspJ2msm6clMCkdW9">Jessica (en-US)</option>
                    <option value="XB0fDUnXU5powFXDhCwa">Charlotte (en-GB)</option>
                    <option value="pFZP5JQG7iQjIQuC4Bku">Lily (en-GB)</option>
                  </select>
                </div>
                <div className="md:col-span-2 text-xs opacity-60">
                  Get a key at elevenlabs.io. Stored only in your browser (localStorage); the app calls ElevenLabs directly from your browser. Uses the eleven_multilingual_v2 model.
                </div>
              </div>
            )}

            {ttsProvider === 'openai' && (
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 flex items-center gap-1.5 mb-1"><Key size={11} /> API key</label>
                  <input type="password" value={openaiKey} onChange={e => setOpenaiKey(e.target.value)} placeholder="sk-..." className="w-full text-sm" />
                </div>
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 block mb-1">Voice</label>
                  <select value={openaiVoice} onChange={e => setOpenaiVoice(e.target.value)} className="w-full text-sm">
                    <option value="alloy">Alloy (neutral)</option>
                    <option value="echo">Echo (male)</option>
                    <option value="fable">Fable (British male)</option>
                    <option value="onyx">Onyx (deep male)</option>
                    <option value="nova">Nova (female)</option>
                    <option value="shimmer">Shimmer (female)</option>
                  </select>
                </div>
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 block mb-1">Speed: <span className="accent font-semibold">{speechRate.toFixed(2)}×</span></label>
                  <input type="range" min="0.5" max="1.5" step="0.05" value={speechRate} onChange={e => setSpeechRate(parseFloat(e.target.value))} className="w-full" />
                </div>
                <div className="md:col-span-2 text-xs opacity-60">
                  Get a key at platform.openai.com. Stored only in your browser (localStorage); the app calls OpenAI directly from your browser. Uses the tts-1-hd model.
                </div>
              </div>
            )}

            {/* Spotify Connection */}
            <div className="mt-6 pt-6" style={{ borderTop: '1px dashed rgba(42,37,32,0.15)' }}>
              <div className="flex items-center justify-between mb-3">
                <div className="mono-font text-xs uppercase tracking-wider opacity-60">Spotify Connection</div>
                {spotifyUser && (
                  <div className="mono-font text-xs flex items-center gap-2">
                    <Check size={12} className="accent" />
                    Connected as <strong>{spotifyUser.display_name || spotifyUser.id}</strong>
                    {spotifyUser.product && <span className="opacity-60">({spotifyUser.product})</span>}
                  </div>
                )}
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="mono-font text-xs uppercase tracking-wider opacity-60 flex items-center gap-1.5 mb-1"><Key size={11} /> Spotify Client ID</label>
                  <input type="text" value={spotifyClientId} onChange={e => setSpotifyClientId(e.target.value)} placeholder="from developer.spotify.com/dashboard" className="w-full text-sm" />
                </div>
                <div className="flex items-end gap-2">
                  {!spotifyToken ? (
                    <button onClick={spotifyConnect} disabled={!spotifyClientId}
                      className="px-4 py-2 mono-font text-xs uppercase tracking-wider accent-bg"
                      style={{ color: '#fdfbf5', borderRadius: '2px' }}>
                      Connect Spotify
                    </button>
                  ) : (
                    <button onClick={spotifyDisconnect}
                      className="px-4 py-2 hairline mono-font text-xs uppercase tracking-wider"
                      style={{ background: 'transparent' }}>
                      Disconnect
                    </button>
                  )}
                  {spotifyToken && (
                    <div className="mono-font text-xs opacity-60">
                      {spotifySdkReady ? '✓ Player ready' : 'Player connecting...'}
                    </div>
                  )}
                </div>
                <div className="md:col-span-2 text-xs opacity-60 leading-relaxed">
                  <strong>Setup (one-time):</strong> at <a href="https://developer.spotify.com/dashboard" target="_blank" rel="noopener" className="underline">developer.spotify.com/dashboard</a>, create an app and add this exact Redirect URI:
                  <code className="mono-font block mt-1 p-2" style={{ background: '#2a2520', color: '#fdfbf5', borderRadius: '2px', wordBreak: 'break-all' }}>
                    {typeof window !== 'undefined' ? window.location.origin + window.location.pathname : ''}
                  </code>
                  Full-track playback requires Spotify Premium. Free accounts can still import playlists and use 30-second track previews.
                </div>
              </div>
            </div>

            <div className="mt-4">
              <button onClick={testVoice} disabled={ttsTestStatus === 'loading'}
                className="flex items-center gap-2 px-3 py-2 hairline mono-font text-xs uppercase tracking-wider"
                style={{ background: 'transparent' }}>
                {ttsTestStatus === 'loading' ? <Loader2 size={12} className="animate-spin" /> : ttsTestStatus === 'ok' ? <Check size={12} className="accent" /> : <Mic size={12} />}
                Test voice
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-6xl mx-auto px-8 py-8">
        {/* PDF + Save/Load toolbar */}
        <section className="mb-8 paper ink-shadow" style={{ borderRadius: '4px', padding: '20px' }}>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {/* PDF upload */}
            <div>
              <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-2">Start from a paper</div>
              <PdfDropZone onFile={handleExamPdfDrop} parsing={pdfParsing} disabled={isCompiling || livePlaying} />
              {pdfDetectionInfo && (
                <div className="mt-2 text-xs flex items-center gap-1.5" style={{ color: '#1db954' }}>
                  <Check size={12} /> Loaded {pdfDetectionInfo.count} extracts
                  {pdfDetectionInfo.marksFound && ' with marks'} from PDF.
                </div>
              )}
            </div>

            {/* Save / Load */}
            <div>
              <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-2">Or resume a saved exam</div>
              <div className="flex gap-2 flex-wrap">
                <button onClick={saveExamConfig}
                  className="flex items-center gap-1.5 px-3 py-2 hairline mono-font text-xs uppercase tracking-wider"
                  style={{ background: 'transparent' }}>
                  <Save size={12} /> Save config
                </button>
                <label className="flex items-center gap-1.5 px-3 py-2 hairline mono-font text-xs uppercase tracking-wider cursor-pointer"
                  style={{ background: 'transparent' }}>
                  <FolderOpen size={12} /> Load config
                  <input type="file" accept=".json,application/json" className="hidden"
                    onChange={e => { loadExamConfig(e.target.files[0]); e.target.value = ''; }} />
                </label>
              </div>
              <div className="mt-2 text-xs opacity-60">
                Saves all settings except uploaded audio files (those need to be re-uploaded). YouTube and Spotify clips are saved with their URLs and timestamps.
              </div>
            </div>
          </div>
        </section>

        <section className="mb-10">
          <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-2">Examination</div>
          <input type="text" value={examTitle} onChange={e => setExamTitle(e.target.value)}
            className="display-font w-full bg-transparent border-none text-4xl font-semibold p-0"
            style={{ borderBottom: '1px solid rgba(42,37,32,0.15)', paddingBottom: '8px' }} />

          <div className="flex flex-wrap gap-8 mt-6 mono-font text-xs uppercase tracking-wider">
            <div><div className="opacity-60 mb-1">Extracts</div><div className="text-lg accent font-semibold">{questions.length}</div></div>
            <div><div className="opacity-60 mb-1">Sources loaded</div><div className="text-lg font-semibold">{filledCount} / {questions.length}</div></div>
            {totalMarks > 0 && (
              <div><div className="opacity-60 mb-1">Total marks</div><div className="text-lg font-semibold flex items-center gap-1.5"><Award size={14} className="accent" /> {totalMarks}</div></div>
            )}
            {youtubeCount > 0 && (
              <div><div className="opacity-60 mb-1">YouTube clips</div><div className="text-lg font-semibold flex items-center gap-1.5"><Youtube size={14} className="accent" /> {youtubeCount}</div></div>
            )}
            {spotifyCount > 0 && (
              <div><div className="opacity-60 mb-1">Spotify clips</div><div className="text-lg font-semibold flex items-center gap-1.5"><Music size={14} style={{ color: '#1db954' }} /> {spotifyCount}</div></div>
            )}
            <div><div className="opacity-60 mb-1">Total runtime</div><div className="text-lg font-semibold">{formatTime(totalDuration)}</div></div>
            <div className="flex items-center gap-3">
              <label className="opacity-60">Reading time (sec)</label>
              <input type="number" value={readingTime} onChange={e => setReadingTime(Math.max(0, parseInt(e.target.value) || 0))} className="w-20 text-sm" />
            </div>
          </div>
        </section>

        <section className="mb-8 paper ink-shadow" style={{ borderRadius: '4px' }}>
          <button onClick={() => setShowScript(!showScript)}
            className="w-full flex items-center justify-between p-5"
            style={{ background: 'transparent', borderBottom: showScript ? '1px solid rgba(42,37,32,0.1)' : 'none' }}>
            <div className="text-left">
              <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-1">Step 01 · Optional</div>
              <h2 className="display-font text-2xl font-semibold italic">Announcement script</h2>
            </div>
            <div className="mono-font text-xs uppercase tracking-wider opacity-60">
              {showScript ? '− Hide' : '+ Edit wording'}
            </div>
          </button>

          {showScript && (
            <div className="p-5 pt-2 space-y-4">
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Opening announcement (played at the very start, before reading time)</label>
                <textarea value={script.opening} onChange={e => setScript({ ...script, opening: e.target.value })} rows={3}
                  className="w-full text-sm italic" style={{ fontFamily: "'Source Serif Pro', serif", resize: 'vertical' }} />
              </div>
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">After reading time (just before the first extract)</label>
                <textarea value={script.postReading} onChange={e => setScript({ ...script, postReading: e.target.value })} rows={2}
                  className="w-full text-sm italic" style={{ fontFamily: "'Source Serif Pro', serif", resize: 'vertical' }} />
              </div>
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">
                  Between plays template — placeholders: <code className="mono-font" style={{ background: '#f0e8d6', padding: '1px 4px', borderRadius: '2px' }}>{'{ord}'}</code> (second/third…) · <code className="mono-font" style={{ background: '#f0e8d6', padding: '1px 4px', borderRadius: '2px' }}>{'{n}'}</code> (2/3…) · <code className="mono-font" style={{ background: '#f0e8d6', padding: '1px 4px', borderRadius: '2px' }}>{'{final}'}</code> (auto-adds " and final" on the last play)
                </label>
                <input type="text" value={script.betweenPlays} onChange={e => setScript({ ...script, betweenPlays: e.target.value })}
                  className="w-full text-sm italic" style={{ fontFamily: "'Source Serif Pro', serif" }} />
                <div className="text-xs opacity-60 mt-1">
                  Example for a 3-play extract → 2nd: <em>"{renderBetweenPlays(script.betweenPlays, 2, false)}"</em> · 3rd: <em>"{renderBetweenPlays(script.betweenPlays, 3, true)}"</em>
                </div>
              </div>
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Closing announcement (after the final extract)</label>
                <textarea value={script.ending} onChange={e => setScript({ ...script, ending: e.target.value })} rows={2}
                  className="w-full text-sm italic" style={{ fontFamily: "'Source Serif Pro', serif", resize: 'vertical' }} />
              </div>
              <div className="flex justify-end pt-2">
                <button onClick={() => setScript(DEFAULT_SCRIPT)}
                  className="mono-font text-xs uppercase tracking-wider opacity-60 hover:opacity-100 underline"
                  style={{ background: 'transparent' }}>
                  Reset to defaults
                </button>
              </div>
            </div>
          )}
        </section>

        {/* Spotify import */}
        {spotifyToken && (
          <section className="mb-8 paper ink-shadow" style={{ borderRadius: '4px', padding: '20px' }}>
            <div className="flex items-baseline justify-between mb-4 flex-wrap gap-2">
              <div>
                <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-1">Spotify · Optional</div>
                <h2 className="display-font text-xl font-semibold italic flex items-center gap-2">
                  <ListMusic size={18} className="accent" /> Import from Spotify
                </h2>
              </div>
              {spotifyPlaylistName && (
                <div className="mono-font text-xs opacity-60">Loaded: <strong>{spotifyPlaylistName}</strong> · {spotifyImportedTracks.length} track{spotifyImportedTracks.length === 1 ? '' : 's'}</div>
              )}
            </div>

            <div className="flex gap-2 flex-wrap items-end">
              <div className="flex-1 min-w-[300px]">
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Playlist or track URL / URI</label>
                <input type="text" value={spotifyPlaylistUrl} onChange={e => setSpotifyPlaylistUrl(e.target.value)}
                  placeholder="https://open.spotify.com/playlist/... or track/..." className="w-full text-sm" />
              </div>
              <button onClick={importSpotifyPlaylist} disabled={spotifyLoading || !spotifyPlaylistUrl}
                className="flex items-center gap-2 px-4 py-2 mono-font text-xs uppercase tracking-wider accent-bg"
                style={{ color: '#fdfbf5', borderRadius: '2px' }}>
                {spotifyLoading ? <Loader2 size={12} className="animate-spin" /> : <ListMusic size={12} />}
                Import playlist
              </button>
              <button onClick={importSpotifyTrack} disabled={spotifyLoading || !spotifyPlaylistUrl}
                className="flex items-center gap-2 px-4 py-2 hairline mono-font text-xs uppercase tracking-wider"
                style={{ background: 'transparent' }}>
                <Music size={12} /> Single track
              </button>
              {spotifyImportedTracks.length > 0 && (
                <button onClick={() => { setSpotifyImportedTracks([]); setSpotifyPlaylistName(''); }}
                  className="px-3 py-2 mono-font text-xs uppercase tracking-wider opacity-60 hover:opacity-100 underline"
                  style={{ background: 'transparent' }}>
                  Clear
                </button>
              )}
            </div>

            {spotifyImportedTracks.length > 0 && (
              <div className="mt-4 space-y-2 max-h-96 overflow-y-auto pr-2" style={{ borderTop: '1px solid rgba(42,37,32,0.1)', paddingTop: '12px' }}>
                {spotifyImportedTracks.map((t, i) => (
                  <SpotifyTrackRow key={t.id} track={t} index={i + 1}
                    questions={questions}
                    onAssign={assignSpotifyTrackToQuestion} />
                ))}
              </div>
            )}
          </section>
        )}

        <section>
          <div className="flex items-baseline justify-between mb-4">
            <div>
              <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-1">Step 02</div>
              <h2 className="display-font text-2xl font-semibold italic">Extracts</h2>
            </div>
            <div className="mono-font text-xs uppercase tracking-wider opacity-60">File · YouTube · Spotify</div>
          </div>

          <div className="space-y-3">
            {questions.map((q, idx) => (
              <QuestionCard key={q.id} q={q} index={idx} totalQuestions={questions.length}
                onFileUpload={handleFileUpload} onYouTubeSet={handleYouTubeSet}
                onSpotifyTrackAdd={handleSpotifyTrackAdd}
                spotifyConnected={!!spotifyToken}
                onClear={clearSource} onUpdate={updateQuestion}
                onPreview={previewQuestion} isPreviewing={previewingId === q.id}
                onMoveUp={idx > 0 ? () => moveQuestion(q.id, 'up') : null}
                onMoveDown={idx < questions.length - 1 ? () => moveQuestion(q.id, 'down') : null}
                onDelete={questions.length > 1 ? () => deleteQuestion(q.id) : null}
                onAddBelow={() => addQuestion(idx)}
                onReorder={moveQuestionTo}
                disabled={isCompiling || livePlaying} />
            ))}

            <button onClick={() => addQuestion()} disabled={isCompiling || livePlaying}
              className="w-full flex items-center justify-center gap-2 py-4 hairline mono-font text-xs uppercase tracking-wider opacity-60 hover:opacity-100"
              style={{ background: 'transparent', borderStyle: 'dashed', borderRadius: '3px' }}>
              <Plus size={14} /> Add extract
            </button>
          </div>
        </section>

        <section className="mt-12 paper ink-shadow" style={{ borderRadius: '4px', padding: '32px' }}>
          <div className="flex items-baseline justify-between mb-6 flex-wrap gap-4">
            <div>
              <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-2">Step 03</div>
              <h2 className="display-font text-3xl font-semibold">Compile & Export</h2>
            </div>
            <label className="flex items-center gap-2 mono-font text-xs uppercase tracking-wider opacity-60 cursor-pointer">
              <input type="checkbox" checked={shortReadingForTesting} onChange={e => setShortReadingForTesting(e.target.checked)} />
              Skip reading time (testing)
            </label>
          </div>

          {isCompiling && (
            <div className="mb-6">
              <div className="mono-font text-xs uppercase tracking-wider mb-2 flex justify-between">
                <span>{compileStatus}</span><span>{compileProgress.toFixed(0)}%</span>
              </div>
              <div className="h-1 hairline rounded overflow-hidden" style={{ background: 'rgba(42,37,32,0.08)' }}>
                <div className="progress-bar h-full transition-all duration-300" style={{ width: `${compileProgress}%` }} />
              </div>
            </div>
          )}

          <div className="flex flex-wrap gap-3">
            <button onClick={playLiveFull} disabled={isCompiling || filledCount === 0}
              className="flex items-center gap-2 px-5 py-3 hairline mono-font text-sm uppercase tracking-wider"
              style={{ background: livePlaying ? '#8b2c1e' : 'transparent', color: livePlaying ? '#fdfbf5' : '#2a2520' }}>
              {livePlaying ? <Pause size={16} /> : <Play size={16} />}
              {livePlaying ? 'Stop preview' : 'Preview full exam (live)'}
            </button>

            <button onClick={compileAudio} disabled={isCompiling || filledCount === 0 || livePlaying}
              className="flex items-center gap-2 px-5 py-3 mono-font text-sm uppercase tracking-wider font-semibold accent-bg"
              style={{ color: '#fdfbf5', borderRadius: '2px' }}>
              <Download size={16} />
              Compile WAV file
            </button>

            {finalAudioUrl && (
              <button onClick={downloadFinal}
                className="flex items-center gap-2 px-5 py-3 hairline mono-font text-sm uppercase tracking-wider"
                style={{ background: 'transparent' }}>
                <FileAudio size={16} />
                Download · {formatTime(finalAudioDuration)}
              </button>
            )}
          </div>

          <div className="mt-6 paper hairline p-4" style={{ borderRadius: '2px', background: '#f0e8d6' }}>
            <div className="mono-font text-xs uppercase tracking-wider opacity-60 mb-2 flex items-center gap-1.5">
              <AlertCircle size={12} /> Source compatibility with WAV export
            </div>
            <ul className="text-sm leading-relaxed space-y-1 ml-4 list-disc">
              <li><strong>Uploaded audio files</strong> — always exported into the WAV.</li>
              <li><strong>ElevenLabs / OpenAI announcements</strong> — fully exported into the WAV when an API key is provided.</li>
              <li><strong>Browser TTS announcements</strong> — cannot be recorded by the browser; the WAV gets a brief marker tone in their place. Use a paid TTS provider to bake voice into the file.</li>
              <li><strong>YouTube clips</strong> — cannot be exported into the WAV (DRM). They <strong>do</strong> play correctly during "Preview full exam (live)". For a fully-exported file, expand any YouTube clip and use the <code className="mono-font">yt-dlp</code> command to extract the clip locally, then upload it as a file.</li>
              <li><strong>Spotify clips</strong> — full-track playback requires Spotify Premium and works only in live preview (DRM). However: if your clip's <em>end time</em> is within the first 30 seconds of a track <em>and</em> Spotify exposes a 30-second preview for that track, the clip <strong>will</strong> be baked into the WAV automatically.</li>
            </ul>
          </div>
        </section>

        <footer className="mt-16 pt-8 text-center mono-font text-xs uppercase tracking-widest opacity-40" style={{ borderTop: '1px solid rgba(42,37,32,0.1)' }}>
          Aural Composer · Prototype v2
        </footer>
      </div>
    </div>
  );
}

function QuestionCard({ q, index, totalQuestions, onFileUpload, onYouTubeSet, onSpotifyTrackAdd, spotifyConnected, onClear, onUpdate, onPreview, isPreviewing, onMoveUp, onMoveDown, onDelete, onAddBelow, onReorder, disabled }) {
  const [dragOver, setDragOver] = useState(false);
  const [mode, setMode] = useState('file');
  const [ytUrl, setYtUrl] = useState('');
  const [ytStart, setYtStart] = useState('');
  const [ytEnd, setYtEnd] = useState('');
  const [spUrl, setSpUrl] = useState('');
  const [spStart, setSpStart] = useState('');
  const [spEnd, setSpEnd] = useState('');
  const fileInputRef = useRef(null);

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith('audio/')) onFileUpload(q.id, file);
  };

  const ytdlpCommand = q.source?.kind === 'youtube'
    ? `yt-dlp -x --audio-format mp3 --download-sections "*${q.source.startStr || formatTime(q.source.start)}-${q.source.endStr || formatTime(q.source.end)}" "${q.source.url}" -o "extract_${q.id}.%(ext)s"`
    : '';

  return (
    <div className="question-card paper ink-shadow"
      style={{ borderRadius: '3px' }}
      draggable={!disabled}
      onDragStart={(e) => {
        e.dataTransfer.setData('text/x-extract-index', String(index));
        e.dataTransfer.effectAllowed = 'move';
      }}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('text/x-extract-index')) {
          e.preventDefault();
          e.dataTransfer.dropEffect = 'move';
        }
      }}
      onDrop={(e) => {
        const fromIdx = e.dataTransfer.getData('text/x-extract-index');
        if (fromIdx !== '') {
          e.preventDefault();
          e.stopPropagation();
          const from = parseInt(fromIdx);
          if (onReorder) onReorder(from, index);
        }
      }}>
      <div className="flex items-stretch">
        <div className="flex flex-col items-center justify-center px-4 py-5 gap-2" style={{ borderRight: '1px solid rgba(42,37,32,0.1)', minWidth: '80px' }}>
          <div className="opacity-30 cursor-grab" title="Drag to reorder">
            <GripVertical size={14} />
          </div>
          <div className="mono-font text-xs uppercase tracking-widest opacity-40">No.</div>
          <div className="display-font text-3xl font-semibold leading-none">{index + 1}</div>
          {q.marks != null && (
            <div className="mono-font text-xs opacity-60 mt-1 flex items-center gap-1">
              <Award size={10} /> {q.marks}
            </div>
          )}
          <div className="flex flex-col gap-0.5 mt-2 opacity-40">
            <button onClick={onMoveUp} disabled={disabled || !onMoveUp} title="Move up"
              className="p-0.5" style={{ background: 'transparent' }}>
              <ChevronUp size={12} />
            </button>
            <button onClick={onMoveDown} disabled={disabled || !onMoveDown} title="Move down"
              className="p-0.5" style={{ background: 'transparent' }}>
              <ChevronDown size={12} />
            </button>
          </div>
        </div>

        <div className="flex-1 p-5">
          <div className="flex items-start justify-between gap-2 mb-3">
            <input type="text" value={q.label} onChange={e => onUpdate(q.id, 'label', e.target.value)} disabled={disabled}
              className="display-font text-xl font-semibold bg-transparent border-none p-0 w-full flex-1"
              style={{ borderBottom: '1px dashed transparent' }}
              onFocus={e => e.target.style.borderBottomColor = 'rgba(42,37,32,0.2)'}
              onBlur={e => e.target.style.borderBottomColor = 'transparent'} />
            <button onClick={() => onPreview(q)} disabled={!q.source || disabled}
              className="flex items-center gap-1 px-3 py-1.5 hairline mono-font text-xs uppercase tracking-wider"
              style={{ background: isPreviewing ? '#8b2c1e' : 'transparent', color: isPreviewing ? '#fdfbf5' : 'inherit' }}
              title="Preview full extract (intro + all plays + between announcements)">
              {isPreviewing ? <Pause size={12} /> : <Play size={12} />}
              Preview
            </button>
            <button onClick={onDelete} disabled={disabled || !onDelete}
              className="p-2 hairline opacity-50 hover:opacity-100"
              style={{ background: 'transparent' }} title="Delete extract">
              <Trash2 size={14} />
            </button>
          </div>

          <div className="mb-4">
            <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Announcement</label>
            <textarea value={q.intro} onChange={e => onUpdate(q.id, 'intro', e.target.value)} disabled={disabled} rows={2}
              className="w-full text-sm italic" style={{ fontFamily: "'Source Serif Pro', serif", resize: 'vertical' }} />
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-4">
            <div>
              <label className="mono-font text-xs uppercase tracking-wider opacity-50 flex items-center gap-1.5 mb-1"><Repeat size={11} /> Plays</label>
              <input type="number" min="1" max="10" value={q.plays} onChange={e => onUpdate(q.id, 'plays', Math.max(1, parseInt(e.target.value) || 1))} disabled={disabled} className="w-full text-sm" />
            </div>
            <div>
              <label className="mono-font text-xs uppercase tracking-wider opacity-50 flex items-center gap-1.5 mb-1"><Clock size={11} /> Gap btwn plays (s)</label>
              <input type="number" min="0" value={q.gapBetweenPlays} onChange={e => onUpdate(q.id, 'gapBetweenPlays', Math.max(0, parseInt(e.target.value) || 0))} disabled={disabled} className="w-full text-sm" />
            </div>
            <div>
              <label className="mono-font text-xs uppercase tracking-wider opacity-50 flex items-center gap-1.5 mb-1"><ChevronRight size={11} /> Gap after (s)</label>
              <input type="number" min="0" value={q.gapAfter} onChange={e => onUpdate(q.id, 'gapAfter', Math.max(0, parseInt(e.target.value) || 0))} disabled={disabled} className="w-full text-sm" />
            </div>
            <div>
              <label className="mono-font text-xs uppercase tracking-wider opacity-50 flex items-center gap-1.5 mb-1"><Award size={11} /> Marks</label>
              <input type="number" min="0" value={q.marks ?? ''} placeholder="—"
                onChange={e => onUpdate(q.id, 'marks', e.target.value === '' ? null : Math.max(0, parseInt(e.target.value) || 0))}
                disabled={disabled} className="w-full text-sm" />
            </div>
          </div>

          {q.source ? (
            <div className="drop-zone hairline has-source p-4" style={{ borderRadius: '2px' }}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3 text-left">
                  {q.source.kind === 'file' && <Music size={18} className="accent" />}
                  {q.source.kind === 'youtube' && <Youtube size={18} className="accent" />}
                  {q.source.kind === 'spotify' && <Music size={18} style={{ color: '#1db954' }} />}
                  <div>
                    <div className="font-semibold text-sm">
                      {q.source.kind === 'file' && q.source.name}
                      {q.source.kind === 'youtube' && `YouTube · ${q.source.videoId}`}
                      {q.source.kind === 'spotify' && (
                        <span>
                          {q.source.name} <span className="opacity-60">— {q.source.artists}</span>
                        </span>
                      )}
                    </div>
                    <div className="mono-font text-xs opacity-70 mt-0.5">
                      {q.source.kind === 'file' && (
                        <>{formatTime(q.source.duration)} · {q.plays}× = {formatTime(q.source.duration * q.plays + q.gapBetweenPlays * (q.plays - 1))}</>
                      )}
                      {q.source.kind === 'youtube' && (
                        <>{q.source.startStr || formatTime(q.source.start)} → {q.source.endStr || formatTime(q.source.end)} · clip {formatTime(q.source.duration)} · {q.plays}× plays</>
                      )}
                      {q.source.kind === 'spotify' && (
                        <>
                          {q.source.startStr} → {q.source.endStr} · clip {formatTime(q.source.duration)} · {q.plays}× plays
                          {q.source.previewUrl && q.source.end <= 30 && <span className="ml-2" style={{ color: '#1db954' }}>✓ fits in 30s preview (WAV export OK)</span>}
                        </>
                      )}
                    </div>
                  </div>
                </div>
                <div className="flex items-center gap-1">
                  {q.source.kind === 'spotify' && q.source.externalUrl && (
                    <a href={q.source.externalUrl} target="_blank" rel="noopener" className="p-2 rounded" style={{ background: 'transparent' }} title="Open in Spotify">
                      <ExternalLink size={14} />
                    </a>
                  )}
                  <button onClick={() => onClear(q.id)} disabled={disabled} className="p-2 rounded" style={{ background: 'transparent' }} title="Remove source">
                    <Trash2 size={14} />
                  </button>
                </div>
              </div>
              {q.source.kind === 'youtube' && (
                <details className="mt-3 pt-3" style={{ borderTop: '1px dashed rgba(42,37,32,0.15)' }}>
                  <summary className="mono-font text-xs uppercase tracking-wider opacity-60">▸ Convert to local file with yt-dlp</summary>
                  <div className="mt-2 flex items-center gap-2">
                    <code className="mono-font text-xs p-2 flex-1 overflow-x-auto" style={{ background: '#2a2520', color: '#fdfbf5', borderRadius: '2px' }}>{ytdlpCommand}</code>
                    <button onClick={() => navigator.clipboard.writeText(ytdlpCommand)} className="p-2 hairline" style={{ background: 'transparent' }} title="Copy">
                      <Copy size={12} />
                    </button>
                  </div>
                  <div className="text-xs opacity-60 mt-2">Run this in your terminal to extract just the clip as MP3, then upload it as a file for full WAV export support.</div>
                </details>
              )}
            </div>
          ) : (
            <div>
              <div className="flex gap-2 mb-2 flex-wrap">
                <button onClick={() => setMode('file')} className={`tab mono-font text-xs uppercase tracking-wider ${mode === 'file' ? 'active' : ''}`} disabled={disabled}>
                  <Upload size={11} className="inline mr-1" /> File
                </button>
                <button onClick={() => setMode('youtube')} className={`tab mono-font text-xs uppercase tracking-wider ${mode === 'youtube' ? 'active' : ''}`} disabled={disabled}>
                  <Youtube size={11} className="inline mr-1" /> YouTube
                </button>
                <button onClick={() => setMode('spotify')} className={`tab mono-font text-xs uppercase tracking-wider ${mode === 'spotify' ? 'active' : ''}`} disabled={disabled}>
                  <Music size={11} className="inline mr-1" /> Spotify
                </button>
              </div>

              {mode === 'file' && (
                <div className="drop-zone hairline p-4 text-center"
                  style={{ borderStyle: 'dashed', borderRadius: '2px', background: dragOver ? '#f0e8d6' : 'transparent' }}
                  onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                  onDragLeave={() => setDragOver(false)}
                  onDrop={handleDrop}>
                  <button onClick={() => fileInputRef.current?.click()} disabled={disabled}
                    className="flex items-center justify-center gap-2 w-full py-3 mono-font text-xs uppercase tracking-wider opacity-60"
                    style={{ background: 'transparent' }}>
                    <Upload size={14} /> Drop audio here or click to choose
                  </button>
                  <input ref={fileInputRef} type="file" accept="audio/*" className="hidden" onChange={e => onFileUpload(q.id, e.target.files[0])} />
                </div>
              )}

              {mode === 'youtube' && (
                <div className="hairline p-4" style={{ borderStyle: 'dashed', borderRadius: '2px' }}>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                    <div className="md:col-span-3">
                      <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">YouTube URL</label>
                      <input type="text" value={ytUrl} onChange={e => setYtUrl(e.target.value)} placeholder="https://www.youtube.com/watch?v=..." className="w-full text-sm" disabled={disabled} />
                    </div>
                    <div>
                      <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Start</label>
                      <input type="text" value={ytStart} onChange={e => setYtStart(e.target.value)} placeholder="0:45 or 1m23s" className="w-full text-sm" disabled={disabled} />
                    </div>
                    <div>
                      <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">End</label>
                      <input type="text" value={ytEnd} onChange={e => setYtEnd(e.target.value)} placeholder="2:15" className="w-full text-sm" disabled={disabled} />
                    </div>
                    <div className="flex items-end">
                      <button onClick={() => {
                        onYouTubeSet(q.id, { url: ytUrl, startStr: ytStart, endStr: ytEnd });
                        setYtUrl(''); setYtStart(''); setYtEnd('');
                      }} disabled={disabled || !ytUrl}
                        className="w-full px-3 py-2 mono-font text-xs uppercase tracking-wider accent-bg"
                        style={{ color: '#fdfbf5', borderRadius: '2px' }}>
                        Add clip
                      </button>
                    </div>
                  </div>
                  <div className="text-xs opacity-50 mt-2">Tip: accepts <code className="mono-font">1:23</code>, <code className="mono-font">83</code>, or <code className="mono-font">1m23s</code> formats.</div>
                </div>
              )}

              {mode === 'spotify' && (
                <div className="hairline p-4" style={{ borderStyle: 'dashed', borderRadius: '2px' }}>
                  {!spotifyConnected ? (
                    <div className="text-sm opacity-70 text-center py-2">
                      Connect to Spotify first (top-right · Voice & API) to add tracks. You can also bulk-import a playlist from the Spotify panel above.
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                      <div className="md:col-span-3">
                        <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Spotify track URL / URI</label>
                        <input type="text" value={spUrl} onChange={e => setSpUrl(e.target.value)} placeholder="https://open.spotify.com/track/..." className="w-full text-sm" disabled={disabled} />
                      </div>
                      <div>
                        <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Start</label>
                        <input type="text" value={spStart} onChange={e => setSpStart(e.target.value)} placeholder="0:00" className="w-full text-sm" disabled={disabled} />
                      </div>
                      <div>
                        <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">End</label>
                        <input type="text" value={spEnd} onChange={e => setSpEnd(e.target.value)} placeholder="2:15 (blank = full track)" className="w-full text-sm" disabled={disabled} />
                      </div>
                      <div className="flex items-end">
                        <button onClick={async () => {
                          await onSpotifyTrackAdd(q.id, spUrl, spStart, spEnd);
                          setSpUrl(''); setSpStart(''); setSpEnd('');
                        }} disabled={disabled || !spUrl}
                          className="w-full px-3 py-2 mono-font text-xs uppercase tracking-wider accent-bg"
                          style={{ color: '#fdfbf5', borderRadius: '2px' }}>
                          Add clip
                        </button>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function audioBufferToWav(buffer) {
  const numChannels = buffer.numberOfChannels;
  const sampleRate = buffer.sampleRate;
  const numSamples = buffer.length;
  const bitDepth = 16;
  const blockAlign = numChannels * bitDepth / 8;
  const byteRate = sampleRate * blockAlign;
  const dataSize = numSamples * blockAlign;
  const arrayBuffer = new ArrayBuffer(44 + dataSize);
  const view = new DataView(arrayBuffer);

  writeString(view, 0, 'RIFF');
  view.setUint32(4, 36 + dataSize, true);
  writeString(view, 8, 'WAVE');
  writeString(view, 12, 'fmt ');
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, numChannels, true);
  view.setUint32(24, sampleRate, true);
  view.setUint32(28, byteRate, true);
  view.setUint16(32, blockAlign, true);
  view.setUint16(34, bitDepth, true);
  writeString(view, 36, 'data');
  view.setUint32(40, dataSize, true);

  const channels = [];
  for (let i = 0; i < numChannels; i++) channels.push(buffer.getChannelData(i));
  let offset = 44;
  for (let i = 0; i < numSamples; i++) {
    for (let ch = 0; ch < numChannels; ch++) {
      let s = Math.max(-1, Math.min(1, channels[ch][i]));
      s = s < 0 ? s * 0x8000 : s * 0x7FFF;
      view.setInt16(offset, s, true);
      offset += 2;
    }
  }
  return new Blob([arrayBuffer], { type: 'audio/wav' });
}

function writeString(view, offset, str) {
  for (let i = 0; i < str.length; i++) view.setUint8(offset + i, str.charCodeAt(i));
}

// ===== PDF drop zone =====
function PdfDropZone({ onFile, parsing, disabled }) {
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  return (
    <div
      className="hairline p-4 text-center"
      style={{
        borderStyle: 'dashed',
        borderRadius: '2px',
        background: dragOver ? '#f0e8d6' : 'transparent',
        opacity: disabled ? 0.5 : 1,
      }}
      onDragOver={e => { e.preventDefault(); if (!disabled) setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={e => {
        e.preventDefault();
        setDragOver(false);
        if (disabled) return;
        const file = e.dataTransfer.files[0];
        if (file) onFile(file);
      }}>
      <button onClick={() => inputRef.current?.click()} disabled={disabled || parsing}
        className="flex items-center justify-center gap-2 w-full py-3 mono-font text-xs uppercase tracking-wider opacity-70"
        style={{ background: 'transparent' }}>
        {parsing ? <><Loader2 size={14} className="animate-spin" /> Reading PDF...</> : <><FileText size={14} /> Drop exam paper PDF here or click</>}
      </button>
      <input ref={inputRef} type="file" accept=".pdf,application/pdf" className="hidden"
        onChange={e => { onFile(e.target.files[0]); e.target.value = ''; }} />
      <div className="text-xs opacity-50 mt-1">Auto-detects extracts, play counts, and marks</div>
    </div>
  );
}

// ===== Spotify track row (in the playlist import staging area) =====
function SpotifyTrackRow({ track, index, questions, onAssign }) {
  const [start, setStart] = useState('0:00');
  const [end, setEnd] = useState('');
  const [selectedQ, setSelectedQ] = useState('');
  const trackDurSec = track.durationMs / 1000;
  const hasPreview = !!track.previewUrl;

  return (
    <div className="flex items-center gap-3 py-2 px-3 hairline" style={{ borderRadius: '2px', background: '#fdfbf5' }}>
      <div className="mono-font text-xs opacity-40 w-6 text-right">{index}</div>
      <div className="flex-1 min-w-0">
        <div className="text-sm font-semibold truncate">{track.name}</div>
        <div className="text-xs opacity-60 truncate">
          {track.artists} · {formatTime(trackDurSec)}
          {hasPreview && <span className="ml-2" style={{ color: '#1db954' }}>· 30s preview</span>}
        </div>
      </div>
      <input type="text" value={start} onChange={e => setStart(e.target.value)} placeholder="start"
        className="text-xs w-16" style={{ padding: '4px 6px' }} />
      <input type="text" value={end} onChange={e => setEnd(e.target.value)} placeholder="end"
        className="text-xs w-16" style={{ padding: '4px 6px' }} />
      <select value={selectedQ} onChange={e => setSelectedQ(e.target.value)}
        className="text-xs" style={{ padding: '4px 6px' }}>
        <option value="">→ Extract...</option>
        {questions.map(q => (
          <option key={q.id} value={q.id}>{q.label}{q.source ? ' (filled)' : ''}</option>
        ))}
      </select>
      <button onClick={() => {
        if (!selectedQ) { alert('Pick an extract first.'); return; }
        onAssign(parseInt(selectedQ), track, start, end || null);
      }} disabled={!selectedQ}
        className="px-2 py-1 mono-font text-xs uppercase tracking-wider accent-bg"
        style={{ color: '#fdfbf5', borderRadius: '2px' }}>
        Assign
      </button>
      {track.externalUrl && (
        <a href={track.externalUrl} target="_blank" rel="noopener" className="opacity-50 hover:opacity-100" title="Open in Spotify">
          <ExternalLink size={12} />
        </a>
      )}
    </div>
  );
}
