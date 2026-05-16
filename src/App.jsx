import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Upload, Play, Pause, Download, Music, Trash2, Clock, Repeat, FileAudio, ChevronRight, Settings, Youtube, Key, Mic, Loader2, Check, AlertCircle, Copy, Link2, ListMusic, ExternalLink, RefreshCw, FileText, Save, FolderOpen, Plus, ChevronUp, ChevronDown, GripVertical, Award, SkipForward, SkipBack, ChevronsRight, Square, Sun, Moon, Menu, X, LogOut, Users, Cloud, CloudOff, ShieldCheck, Mail, HelpCircle, Sparkles } from 'lucide-react';
import { supabase, getMyProfile, listExams, saveExam as supaSaveExam, deleteExam as supaDeleteExam, listAllProfiles, setUserApproved } from './supabaseClient';

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

function computeWaveformPeaks(audioBuffer, numBars) {
  // Mix down all channels to mono peaks
  const channels = audioBuffer.numberOfChannels;
  const data = audioBuffer.getChannelData(0);
  // Mix in other channels by averaging
  let mixed = data;
  if (channels > 1) {
    mixed = new Float32Array(data.length);
    for (let ch = 0; ch < channels; ch++) {
      const cd = audioBuffer.getChannelData(ch);
      for (let i = 0; i < cd.length; i++) mixed[i] += cd[i] / channels;
    }
  }
  const samplesPerBar = Math.max(1, Math.floor(mixed.length / numBars));
  const peaks = new Float32Array(numBars);
  for (let b = 0; b < numBars; b++) {
    let max = 0;
    const start = b * samplesPerBar;
    const end = Math.min(start + samplesPerBar, mixed.length);
    for (let i = start; i < end; i++) {
      const v = Math.abs(mixed[i]);
      if (v > max) max = v;
    }
    peaks[b] = max;
  }
  // Return as plain array for easy JSON serialisation later if needed
  return Array.from(peaks);
}

// Compute RMS (root mean square) loudness of an audio buffer.
// Returns a single value 0..1 representing the average signal level.
function computeRms(audioBuffer) {
  const channels = audioBuffer.numberOfChannels;
  let sumSquares = 0;
  let count = 0;
  for (let ch = 0; ch < channels; ch++) {
    const data = audioBuffer.getChannelData(ch);
    // Sample every 10th value for speed — plenty for an average
    for (let i = 0; i < data.length; i += 10) {
      sumSquares += data[i] * data[i];
      count++;
    }
  }
  if (count === 0) return 0;
  return Math.sqrt(sumSquares / count);
}

const kbdStyle = {
  background: 'var(--surface-elev)',
  border: '0.5px solid var(--border)',
  padding: '2px 6px',
  borderRadius: '4px',
  fontSize: '11px',
  color: 'var(--text)',
};

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
  // Match either web URL (open.spotify.com/playlist/ID, with optional ?query) or URI (spotify:playlist:ID).
  // Spotify IDs are 22 chars base62 in current usage, but the API also accepts 21- and 23-char ids historically — accept any length 20-24 to be safe.
  const re = new RegExp(`${kind}[/:]([a-zA-Z0-9]{20,24})`, 'i');
  const m = url.match(re);
  if (m) return m[1];
  // Bare ID (assume current length range)
  const trimmed = url.trim();
  if (/^[a-zA-Z0-9]{20,24}$/.test(trimmed)) return trimmed;
  return null;
}

// ===== Spotify PKCE OAuth =====
// Generate cryptographically random string for PKCE verifier
function generateRandomString(length) {
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  const arr = new Uint8Array(length);
  window.crypto.getRandomValues(arr);
  return Array.from(arr).map(x => possible[x % possible.length]).join('');
}

// SHA256 hash + base64url-encode (for PKCE challenge)
async function sha256Base64Url(input) {
  const data = new TextEncoder().encode(input);
  const hash = await window.crypto.subtle.digest('SHA-256', data);
  const bytes = new Uint8Array(hash);
  let binary = '';
  for (const b of bytes) binary += String.fromCharCode(b);
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

async function startSpotifyPkceFlow(clientId, redirectUri) {
  const verifier = generateRandomString(64);
  const challenge = await sha256Base64Url(verifier);
  sessionStorage.setItem('aural_spotify_pkce_verifier', verifier);
  sessionStorage.setItem('aural_spotify_pkce_client_id', clientId);
  sessionStorage.setItem('aural_spotify_pkce_redirect', redirectUri);

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
    response_type: 'code',
    redirect_uri: redirectUri,
    scope: scopes,
    code_challenge_method: 'S256',
    code_challenge: challenge,
  });
  window.location.href = `https://accounts.spotify.com/authorize?${params.toString()}`;
}

// Exchange the code returned by Spotify for an access token
async function exchangeSpotifyCode(code) {
  const verifier = sessionStorage.getItem('aural_spotify_pkce_verifier');
  const clientId = sessionStorage.getItem('aural_spotify_pkce_client_id');
  const redirectUri = sessionStorage.getItem('aural_spotify_pkce_redirect');
  if (!verifier || !clientId || !redirectUri) {
    throw new Error('PKCE state missing (sessionStorage cleared?). Please reconnect.');
  }
  const params = new URLSearchParams({
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri,
    client_id: clientId,
    code_verifier: verifier,
  });
  const res = await fetch('https://accounts.spotify.com/api/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Token exchange failed (${res.status}): ${t.slice(0, 200)}`);
  }
  const data = await res.json();
  // Clean up sessionStorage
  sessionStorage.removeItem('aural_spotify_pkce_verifier');
  return {
    token: data.access_token,
    refreshToken: data.refresh_token,
    expiresAt: Date.now() + (data.expires_in || 3600) * 1000,
    clientId,
  };
}

// Check the current URL for a Spotify auth code (after redirect)
function getSpotifyAuthCodeFromUrl() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get('code');
  const error = params.get('error');
  if (!code && !error) return null;
  // Clean the URL to remove the code/error params
  window.history.replaceState({}, document.title, window.location.pathname);
  if (error) return { error };
  return { code };
}

// Refresh an expired token using its refresh token
async function refreshSpotifyToken(refreshToken, clientId) {
  const params = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    client_id: clientId,
  });
  const res = await fetch('https://accounts.spotify.com/api/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString(),
  });
  if (!res.ok) throw new Error(`Token refresh failed: ${res.status}`);
  const data = await res.json();
  return {
    token: data.access_token,
    refreshToken: data.refresh_token || refreshToken,
    expiresAt: Date.now() + (data.expires_in || 3600) * 1000,
    clientId,
  };
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
  // ===== Auth state =====
  const [authLoading, setAuthLoading] = useState(true);
  const [session, setSession] = useState(null);
  const [myProfile, setMyProfile] = useState(null);
  const [cloudExams, setCloudExams] = useState([]);
  const [cloudExamsLoading, setCloudExamsLoading] = useState(false);
  const [showAdminPanel, setShowAdminPanel] = useState(false);
  const [showHelpPanel, setShowHelpPanel] = useState(false);
  const [showWelcomeBanner, setShowWelcomeBanner] = useState(() => {
    return localStorage.getItem('aural_welcome_dismissed') !== '1';
  });
  const [showQuickStart, setShowQuickStart] = useState(() => {
    return localStorage.getItem('aural_quickstart_dismissed') !== '1';
  });
  const dismissWelcome = () => {
    setShowWelcomeBanner(false);
    localStorage.setItem('aural_welcome_dismissed', '1');
  };
  const dismissQuickStart = () => {
    setShowQuickStart(false);
    localStorage.setItem('aural_quickstart_dismissed', '1');
  };

  // Set up auth listener on mount
  useEffect(() => {
    let unsubscribed = false;
    supabase.auth.getSession().then(({ data: { session } }) => {
      if (unsubscribed) return;
      setSession(session);
      setAuthLoading(false);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      if (unsubscribed) return;
      setSession(session);
    });
    return () => { unsubscribed = true; subscription?.unsubscribe?.(); };
  }, []);

  // When session changes, fetch profile + cloud exams
  useEffect(() => {
    if (!session) {
      setMyProfile(null);
      setCloudExams([]);
      return;
    }
    (async () => {
      const profile = await getMyProfile();
      setMyProfile(profile);
      if (profile?.approved) {
        setCloudExamsLoading(true);
        try {
          const exams = await listExams();
          setCloudExams(exams);
        } catch (err) {
          console.warn('Could not load cloud exams:', err.message);
        } finally {
          setCloudExamsLoading(false);
        }
      }
    })();
  }, [session]);

  const reloadCloudExams = async () => {
    if (!myProfile?.approved) return;
    setCloudExamsLoading(true);
    try {
      const exams = await listExams();
      setCloudExams(exams);
    } catch (err) {
      console.warn('Reload failed:', err.message);
    } finally {
      setCloudExamsLoading(false);
    }
  };

  const signOut = async () => {
    if (!confirm('Sign out?')) return;
    await supabase.auth.signOut();
  };

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
  const [finalAudioFormat, setFinalAudioFormat] = useState(null); // 'wav' | 'mp3' | 'ogg' | 'webm'
  const [finalAudioSize, setFinalAudioSize] = useState(0);
  const [outputFormat, setOutputFormat] = useState(() => localStorage.getItem('aural_output_format') || 'mp3');
  const [mp3Bitrate, setMp3Bitrate] = useState(() => parseInt(localStorage.getItem('aural_mp3_bitrate') || '192', 10));
  const [normaliseLoudness, setNormaliseLoudness] = useState(() => localStorage.getItem('aural_normalise') !== 'false');
  const [crossfadeMs, setCrossfadeMs] = useState(() => parseInt(localStorage.getItem('aural_crossfade') || '150', 10));
  useEffect(() => { localStorage.setItem('aural_output_format', outputFormat); }, [outputFormat]);
  useEffect(() => { localStorage.setItem('aural_mp3_bitrate', String(mp3Bitrate)); }, [mp3Bitrate]);
  useEffect(() => { localStorage.setItem('aural_normalise', String(normaliseLoudness)); }, [normaliseLoudness]);
  useEffect(() => { localStorage.setItem('aural_crossfade', String(crossfadeMs)); }, [crossfadeMs]);
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

  // Capture Spotify OAuth callback on mount (PKCE code → token exchange)
  useEffect(() => {
    const callback = getSpotifyAuthCodeFromUrl();
    if (!callback) return;
    if (callback.error) {
      alert(`Spotify authentication was cancelled: ${callback.error}`);
      sessionStorage.removeItem('aural_spotify_pkce_verifier');
      return;
    }
    (async () => {
      try {
        const tok = await exchangeSpotifyCode(callback.code);
        setSpotifyToken(tok);
        localStorage.setItem('aural_spotify_token', JSON.stringify(tok));
      } catch (err) {
        console.error(err);
        alert(`Could not complete Spotify connection: ${err.message}`);
      }
    })();
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
      const peaks = computeWaveformPeaks(audioBuffer, 60);
      const rms = computeRms(audioBuffer);
      setSource(questionId, {
        kind: 'file',
        name: file.name,
        buffer: audioBuffer,
        duration: audioBuffer.duration,
        peaks,
        rms,
      });
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
  const spotifyConnect = async () => {
    if (!spotifyClientId) {
      alert('Please enter your Spotify Client ID first (Voice & API → Spotify).');
      return;
    }
    const redirectUri = window.location.origin + window.location.pathname;
    try {
      await startSpotifyPkceFlow(spotifyClientId, redirectUri);
      // Note: this navigates away, so code after this won't execute.
    } catch (err) {
      alert(`Could not start Spotify login: ${err.message}`);
    }
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

    // Pre-emptively refresh if within 60s of expiry
    let tokenToUse = spotifyToken;
    if (tokenToUse.refreshToken && tokenToUse.expiresAt && tokenToUse.expiresAt - Date.now() < 60_000) {
      try {
        const fresh = await refreshSpotifyToken(tokenToUse.refreshToken, tokenToUse.clientId);
        setSpotifyToken(fresh);
        localStorage.setItem('aural_spotify_token', JSON.stringify(fresh));
        tokenToUse = fresh;
      } catch (e) { console.warn('Token refresh failed:', e); }
    }

    let res = await fetch(url, { headers: { Authorization: `Bearer ${tokenToUse.token}` } });
    if (res.status === 401 && tokenToUse.refreshToken) {
      // Retry once with a refreshed token
      try {
        const fresh = await refreshSpotifyToken(tokenToUse.refreshToken, tokenToUse.clientId);
        setSpotifyToken(fresh);
        localStorage.setItem('aural_spotify_token', JSON.stringify(fresh));
        res = await fetch(url, { headers: { Authorization: `Bearer ${fresh.token}` } });
      } catch (e) {
        spotifyDisconnect();
        throw new Error('Spotify session expired. Please reconnect.');
      }
    }
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
      const meta = await spotifyFetch(`https://api.spotify.com/v1/playlists/${playlistId}?fields=name,description`);
      setSpotifyPlaylistName(meta.name);

      let allTracks = [];
      // Spotify renamed /tracks → /items in Feb 2026. Response field renames: items[].track → items[].item
      // We support both shapes for safety.
      let next = `https://api.spotify.com/v1/playlists/${playlistId}/items?limit=100`;
      while (next) {
        const page = await spotifyFetch(next);
        const rows = page.items || [];
        for (const row of rows) {
          // New format: row.item. Old format: row.track. Some responses include both.
          const t = row.item || row.track;
          if (t && t.id) {
            allTracks.push({
              id: t.id,
              name: t.name,
              artists: (t.artists || []).map(a => a.name).join(', '),
              durationMs: t.duration_ms,
              previewUrl: t.preview_url,
              uri: t.uri,
              externalUrl: t.external_urls?.spotify,
            });
          }
        }
        next = page.next;
      }
      if (allTracks.length === 0) {
        alert('Playlist loaded but contained no playable tracks (or Spotify returned an unexpected response format).');
      }
      setSpotifyImportedTracks(allTracks);
    } catch (err) {
      // Surface 403 with a friendlier explanation
      let msg = err.message;
      if (msg.includes('403')) {
        msg = 'Spotify refused this request (403).\n\nThis often means your app needs Extended Quota Mode for playlist access, or the playlist isn\'t accessible to your account. See the dev console for details.';
      }
      alert(`Failed to import playlist: ${msg}`);
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

  // Load and decode a Spotify preview URL into an AudioBuffer (used for audio export when clip fits in 30s preview)
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
    const config = buildConfig();
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
      applyConfig(config);
    } catch (err) {
      alert(`Could not load config: ${err.message}`);
    }
  };

  // Apply a config object to current state
  const applyConfig = (config) => {
    setExamTitle(config.examTitle || '');
    setReadingTime(config.readingTime || 300);
    if (config.script) setScript(config.script);
    setQuestions(config.questions.map(q => ({
      ...q,
      source: q.source && q.source.kind !== 'file' ? q.source : null,
    })));
  };

  // Build a config object from current state
  const buildConfig = () => ({
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
      source: q.source && q.source.kind !== 'file' ? q.source : null,
    })),
  });

  // ===== Browser-stored exam library =====
  const SAVED_EXAMS_KEY = 'aural_saved_exams';

  const [savedExams, setSavedExams] = useState(() => {
    try {
      const raw = localStorage.getItem(SAVED_EXAMS_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) { return []; }
  });

  const persistSavedExams = (list) => {
    setSavedExams(list);
    try { localStorage.setItem(SAVED_EXAMS_KEY, JSON.stringify(list)); } catch (e) {
      alert('Could not save: browser storage may be full.');
    }
  };

  const saveCurrentExam = async () => {
    const defaultName = examTitle || `Exam ${new Date().toLocaleDateString()}`;
    const name = prompt('Save this exam as:', defaultName);
    if (!name) return;
    const config = buildConfig();

    // Prefer cloud if signed in & approved
    if (myProfile?.approved) {
      try {
        await supaSaveExam({ name, config, shared_with_all: false });
        await reloadCloudExams();
        return;
      } catch (err) {
        alert(`Cloud save failed: ${err.message}\n\nFalling back to browser storage.`);
      }
    }

    // Local fallback (or only mode if signed out)
    const entry = {
      id: `exam_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      name,
      savedAt: new Date().toISOString(),
      config,
    };
    const filtered = savedExams.filter(x => x.name !== name);
    persistSavedExams([entry, ...filtered]);
  };

  const updateExistingExam = async (entry) => {
    if (!confirm(`Update "${entry.name}" with current settings?`)) return;
    if (entry.kind === 'cloud') {
      try {
        await supaSaveExam({ id: entry.id, name: entry.name, config: buildConfig(), shared_with_all: entry.shared_with_all });
        await reloadCloudExams();
      } catch (err) {
        alert(`Could not update cloud exam: ${err.message}`);
      }
    } else {
      const idx = savedExams.findIndex(x => x.id === entry.id);
      if (idx === -1) return;
      const updated = [...savedExams];
      updated[idx] = { ...updated[idx], config: buildConfig(), savedAt: new Date().toISOString() };
      persistSavedExams(updated);
    }
  };

  const loadSavedExam = (entry) => {
    const anyFilled = questions.some(q => q.source);
    if (anyFilled && !confirm(`Load "${entry.name}"? Uploaded audio in the current exam will be cleared (other sources are kept).`)) return;
    applyConfig(entry.config);
  };

  const renameSavedExam = async (entry) => {
    const newName = prompt('Rename to:', entry.name);
    if (!newName || newName === entry.name) return;
    if (entry.kind === 'cloud') {
      try {
        await supaSaveExam({ id: entry.id, name: newName, config: entry.config, shared_with_all: entry.shared_with_all });
        await reloadCloudExams();
      } catch (err) {
        alert(`Could not rename: ${err.message}`);
      }
    } else {
      persistSavedExams(savedExams.map(x => x.id === entry.id ? { ...x, name: newName } : x));
    }
  };

  const deleteSavedExam = async (entry) => {
    if (!confirm(`Delete "${entry.name}"? This cannot be undone.`)) return;
    if (entry.kind === 'cloud') {
      try {
        await supaDeleteExam(entry.id);
        await reloadCloudExams();
      } catch (err) {
        alert(`Could not delete: ${err.message}`);
      }
    } else {
      persistSavedExams(savedExams.filter(x => x.id !== entry.id));
    }
  };

  const toggleShareWithAll = async (entry) => {
    if (entry.kind !== 'cloud') {
      alert('Sharing is only available for cloud-saved exams. Save this exam first.');
      return;
    }
    try {
      await supaSaveExam({ id: entry.id, name: entry.name, config: entry.config, shared_with_all: !entry.shared_with_all });
      await reloadCloudExams();
    } catch (err) {
      alert(`Could not change sharing: ${err.message}`);
    }
  };

  const newBlankExam = () => {
    if (!confirm('Start a fresh exam? Current settings will be cleared (you can save first if needed).')) return;
    setExamTitle('Untitled exam');
    setQuestions(DEFAULT_QUESTIONS.map(q => ({ ...q, source: null })));
    setScript(DEFAULT_SCRIPT);
    setReadingTime(300);
  };

  // ===== Share-via-URL =====
  const [shareToastMessage, setShareToastMessage] = useState(null);

  const shareAsUrl = async () => {
    try {
      const LZString = await loadLzString();
      const config = buildConfig();
      const json = JSON.stringify(config);
      const compressed = LZString.compressToEncodedURIComponent(json);
      const url = `${window.location.origin}${window.location.pathname}#exam=${compressed}`;
      // Try clipboard
      try {
        await navigator.clipboard.writeText(url);
        setShareToastMessage(`Link copied · ${(url.length / 1024).toFixed(1)} KB · ${config.questions.filter(q => q.source && q.source.kind !== 'file').length} non-file sources kept, audio files not included`);
      } catch (e) {
        // Fallback: prompt
        prompt('Copy this share link (audio files not included):', url);
      }
      setTimeout(() => setShareToastMessage(null), 6000);
    } catch (err) {
      alert(`Could not create share link: ${err.message}`);
    }
  };

  // On mount, check URL hash for a shared config
  useEffect(() => {
    const hash = window.location.hash;
    if (!hash.startsWith('#exam=')) return;
    const encoded = hash.slice('#exam='.length);
    (async () => {
      try {
        const LZString = await loadLzString();
        const json = LZString.decompressFromEncodedURIComponent(encoded);
        if (!json) throw new Error('Invalid share link');
        const config = JSON.parse(json);
        if (!config.questions || !Array.isArray(config.questions)) throw new Error('Invalid config in link');
        const proceed = confirm(`Load shared exam "${config.examTitle || 'Untitled'}"?\n\n${config.questions.length} extracts. Note: any uploaded audio files in the original exam are NOT included (only YouTube/Spotify references). You'll need to re-upload any local audio.\n\nThis will replace your current setup.`);
        if (proceed) {
          applyConfig(config);
          // Clear hash so refreshing doesn't re-trigger
          window.history.replaceState({}, document.title, window.location.pathname);
          setShareToastMessage(`Loaded shared exam · click Save in the sidebar to keep it`);
          setTimeout(() => setShareToastMessage(null), 8000);
        } else {
          window.history.replaceState({}, document.title, window.location.pathname);
        }
      } catch (err) {
        alert(`Could not load shared link: ${err.message}`);
      }
    })();
    // Run once on mount
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [mobileDrawerOpen, setMobileDrawerOpen] = useState(false);
  const [theme, setTheme] = useState(() => {
    const saved = localStorage.getItem('aural_theme');
    if (saved === 'light' || saved === 'dark') return saved;
    if (typeof window !== 'undefined' && window.matchMedia('(prefers-color-scheme: light)').matches) return 'light';
    return 'dark';
  });

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('aural_theme', theme);
  }, [theme]);

  // Listen for system theme changes (only if user hasn't set a manual preference recently)
  useEffect(() => {
    const mq = window.matchMedia('(prefers-color-scheme: light)');
    const onChange = (e) => {
      // Only auto-switch if the user hasn't toggled within the last 24 hours
      const lastManual = parseInt(localStorage.getItem('aural_theme_manual_at') || '0', 10);
      if (Date.now() - lastManual > 24 * 60 * 60 * 1000) {
        setTheme(e.matches ? 'light' : 'dark');
      }
    };
    mq.addEventListener?.('change', onChange);
    return () => mq.removeEventListener?.('change', onChange);
  }, []);

  const toggleTheme = () => {
    setTheme(prev => prev === 'dark' ? 'light' : 'dark');
    localStorage.setItem('aural_theme_manual_at', String(Date.now()));
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
            const trimStart = q.source.trimStart || 0;
            const trimEnd = q.source.trimEnd != null ? q.source.trimEnd : q.source.buffer.duration;
            const playDur = trimEnd - trimStart;
            timeline.push({
              type: 'audio',
              buffer: q.source.buffer,
              bufferOffset: trimStart,
              playDuration: playDur,
              start: cursor,
              duration: playDur,
              questionId: q.id,
              playNumber: i + 1,
            });
            cursor += playDur;
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
            // Thinking time FIRST (silence), then the announcement as a cue for the next play
            const silenceBefore = Math.max(0, q.gapBetweenPlays - announceDur);
            addSilence(silenceBefore, 'Thinking time');
            timeline.push({ type: 'tts', text: announce, start: cursor, duration: announceDur });
            cursor += announceDur;
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
            const trimStart = q.source.trimStart || 0;
            const trimEnd = q.source.trimEnd != null ? q.source.trimEnd : q.source.buffer.duration;
            src.start(0, trimStart, trimEnd - trimStart);
            previewStopRef.current = { stop: () => { try { src.stop(); } catch(e){} } };
            src.onended = resolve;
          });
        } else if (q.source?.kind === 'youtube') {
          await playYouTubeSegment(q.source.videoId, q.source.start, q.source.end);
        } else if (q.source?.kind === 'spotify') {
          await playSpotifySegment(q.source.uri, q.source.start, q.source.end);
        }
        // Between-plays announcement (not after last play): thinking time first, then announcement as cue
        if (i < q.plays - 1) {
          const upcoming = i + 2;
          const announce = renderBetweenPlays(script.betweenPlays, upcoming, upcoming === q.plays);
          // Short pause to simulate thinking time, capped at 5s for preview brevity
          const pauseMs = Math.min(q.gapBetweenPlays * 1000, 5000);
          await new Promise(r => setTimeout(r, pauseMs));
          await speakLive(announce);
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

  // ===== Live playback state =====
  const [livePaused, setLivePaused] = useState(false);
  const [liveItemIndex, setLiveItemIndex] = useState(0);
  const [liveTotalItems, setLiveTotalItems] = useState(0);
  const [liveCurrentLabel, setLiveCurrentLabel] = useState('');
  const livePlaybackRef = useRef({
    timeline: [],
    cursor: 0,
    abortCurrent: null, // function: aborts current item, takes 'skip' | 'pause'
    paused: false,
    skipRequest: null, // null | 'next-extract' | 'prev-extract' | 'skip-item'
    stopped: false,
    runId: 0,
  });

  const describeItem = (item, timeline, idx) => {
    if (item.type === 'tts') return 'Announcement';
    if (item.type === 'silence') return item.label || 'Silence';
    if (item.type === 'audio' || item.type === 'youtube' || item.type === 'spotify') {
      const q = questions.find(x => x.id === item.questionId);
      const label = q?.label || `Extract ${item.questionId}`;
      return `${label} · play ${item.playNumber} of ${q?.plays || '?'}`;
    }
    return 'Playing...';
  };

  const findExtractBoundary = (timeline, currentIdx, direction) => {
    // Find next/prev item where type is audio/youtube/spotify with playNumber === 1 (start of an extract)
    if (direction === 'next') {
      for (let i = currentIdx + 1; i < timeline.length; i++) {
        const t = timeline[i];
        if ((t.type === 'audio' || t.type === 'youtube' || t.type === 'spotify') && t.playNumber === 1) return i;
      }
      return timeline.length; // end
    } else {
      // Previous: find the most recent extract start at or before currentIdx, then go one more back
      let lastStart = -1;
      for (let i = 0; i < currentIdx; i++) {
        const t = timeline[i];
        if ((t.type === 'audio' || t.type === 'youtube' || t.type === 'spotify') && t.playNumber === 1) {
          lastStart = i;
        }
      }
      // If we're currently AT the start of an extract, go back to the previous one's start
      // To do this we need the start before lastStart
      let prevStart = -1;
      for (let i = 0; i < lastStart; i++) {
        const t = timeline[i];
        if ((t.type === 'audio' || t.type === 'youtube' || t.type === 'spotify') && t.playNumber === 1) {
          prevStart = i;
        }
      }
      if (currentIdx > lastStart && lastStart >= 0) return lastStart;
      if (prevStart >= 0) return prevStart;
      return 0; // jump to start of timeline
    }
  };

  const playLiveFull = async () => {
    if (livePlaying) {
      // Stop entirely
      livePlaybackRef.current.stopped = true;
      livePlaybackRef.current.abortCurrent?.('stop');
      stopAll();
      setLivePaused(false);
      return;
    }

    const { timeline } = buildTimeline();
    const myRunId = (livePlaybackRef.current.runId || 0) + 1;
    livePlaybackRef.current = {
      timeline,
      cursor: 0,
      abortCurrent: null,
      paused: false,
      skipRequest: null,
      stopped: false,
      runId: myRunId,
    };
    setLivePlaying(true);
    setLivePaused(false);
    setLiveTotalItems(timeline.length);

    const ctx = getAudioContext();
    if (ctx.state === 'suspended') await ctx.resume();

    while (livePlaybackRef.current.cursor < timeline.length) {
      if (livePlaybackRef.current.stopped || livePlaybackRef.current.runId !== myRunId) break;

      // Wait while paused
      while (livePlaybackRef.current.paused) {
        if (livePlaybackRef.current.stopped || livePlaybackRef.current.runId !== myRunId) break;
        await new Promise(r => setTimeout(r, 100));
      }
      if (livePlaybackRef.current.stopped || livePlaybackRef.current.runId !== myRunId) break;

      // Handle skip requests
      if (livePlaybackRef.current.skipRequest === 'next-extract') {
        livePlaybackRef.current.cursor = findExtractBoundary(timeline, livePlaybackRef.current.cursor, 'next');
        livePlaybackRef.current.skipRequest = null;
        continue;
      } else if (livePlaybackRef.current.skipRequest === 'prev-extract') {
        livePlaybackRef.current.cursor = findExtractBoundary(timeline, livePlaybackRef.current.cursor, 'prev');
        livePlaybackRef.current.skipRequest = null;
        continue;
      } else if (livePlaybackRef.current.skipRequest === 'skip-item') {
        livePlaybackRef.current.cursor++;
        livePlaybackRef.current.skipRequest = null;
        continue;
      }

      const idx = livePlaybackRef.current.cursor;
      const item = timeline[idx];
      setLiveItemIndex(idx);
      setLiveCurrentLabel(describeItem(item, timeline, idx));

      try {
        await playLiveItem(item, ctx);
      } catch (e) {
        console.warn('Live item error:', e);
      }

      // If a skip was requested DURING the item, the abort handler set skipRequest
      // and the loop will pick it up. Otherwise advance.
      if (!livePlaybackRef.current.skipRequest && !livePlaybackRef.current.stopped) {
        livePlaybackRef.current.cursor++;
      }
    }

    setLivePlaying(false);
    setLivePaused(false);
    setLiveCurrentLabel('');
  };

  // Play a single timeline item, returns when done or aborted
  const playLiveItem = (item, ctx) => {
    return new Promise((resolve) => {
      let done = false;
      const finish = () => { if (!done) { done = true; livePlaybackRef.current.abortCurrent = null; resolve(); } };

      if (item.type === 'tts') {
        let utter = null;
        const abort = () => { try { window.speechSynthesis.cancel(); } catch (e) {} finish(); };
        livePlaybackRef.current.abortCurrent = abort;
        // Pause-aware: speakLive returns a promise that resolves when done
        speakLive(item.text).then(finish).catch(finish);
      } else if (item.type === 'silence') {
        const dur = item.duration * 1000;
        let elapsed = 0;
        let timeoutId = null;
        const tick = () => {
          if (done) return;
          if (!livePlaybackRef.current.paused) {
            elapsed += 100;
          }
          if (elapsed >= dur) { finish(); return; }
          timeoutId = setTimeout(tick, 100);
        };
        const abort = () => { clearTimeout(timeoutId); finish(); };
        livePlaybackRef.current.abortCurrent = abort;
        timeoutId = setTimeout(tick, 100);
      } else if (item.type === 'audio') {
        const src = ctx.createBufferSource();
        src.buffer = item.buffer;
        src.connect(ctx.destination);
        let stoppedManually = false;
        const abort = () => { stoppedManually = true; try { src.stop(); } catch (e) {} finish(); };
        livePlaybackRef.current.abortCurrent = abort;
        src.onended = () => finish();
        if (item.bufferOffset != null) {
          src.start(0, item.bufferOffset, item.playDuration);
        } else {
          src.start();
        }
      } else if (item.type === 'youtube') {
        const abort = () => {
          if (ytPlayerRef.current) {
            try { ytPlayerRef.current.stopVideo(); } catch (e) {}
          }
          finish();
        };
        livePlaybackRef.current.abortCurrent = abort;
        playYouTubeSegment(item.videoId, item.startSec, item.endSec).then(finish).catch(finish);
      } else if (item.type === 'spotify') {
        const abort = () => {
          if (spotifyPlayerRef.current) {
            try { spotifyPlayerRef.current.pause(); } catch (e) {}
          }
          finish();
        };
        livePlaybackRef.current.abortCurrent = abort;
        playSpotifySegment(item.uri, item.startSec, item.endSec).then(finish).catch(finish);
      } else {
        finish();
      }
    });
  };

  const pauseLive = () => {
    if (!livePlaying) return;
    livePlaybackRef.current.paused = true;
    setLivePaused(true);
    // Pause underlying players
    try { window.speechSynthesis.pause(); } catch (e) {}
    if (ytPlayerRef.current) { try { ytPlayerRef.current.pauseVideo(); } catch (e) {} }
    if (spotifyPlayerRef.current) { try { spotifyPlayerRef.current.pause(); } catch (e) {} }
    // Audio sources can't be paused; they keep playing until done. Acceptable for short clips.
  };

  const resumeLive = () => {
    if (!livePlaying) return;
    livePlaybackRef.current.paused = false;
    setLivePaused(false);
    try { window.speechSynthesis.resume(); } catch (e) {}
    if (ytPlayerRef.current) { try { ytPlayerRef.current.playVideo(); } catch (e) {} }
    if (spotifyPlayerRef.current) { try { spotifyPlayerRef.current.resume(); } catch (e) {} }
  };

  const skipToNextExtract = () => {
    if (!livePlaying) return;
    livePlaybackRef.current.skipRequest = 'next-extract';
    livePlaybackRef.current.abortCurrent?.('skip');
  };

  const skipToPrevExtract = () => {
    if (!livePlaying) return;
    livePlaybackRef.current.skipRequest = 'prev-extract';
    livePlaybackRef.current.abortCurrent?.('skip');
  };

  const skipCurrentItem = () => {
    if (!livePlaying) return;
    livePlaybackRef.current.skipRequest = 'skip-item';
    livePlaybackRef.current.abortCurrent?.('skip');
  };

  const compileAudio = async () => {
    const filled = questions.filter(q => q.source);
    if (filled.length === 0) { alert('Add audio to at least one extract before compiling.'); return; }

    const hasYouTube = filled.some(q => q.source.kind === 'youtube');
    const hasSpotify = filled.some(q => q.source.kind === 'spotify');
    const usingPremiumTTS = (ttsProvider === 'eleven' && elevenKey) || (ttsProvider === 'openai' && openaiKey);

    // Collect warnings about what won't be in the WAV
    const warnings = [];
    if (!usingPremiumTTS) {
      warnings.push('• Announcements will be marker tones (browser TTS cannot be recorded into the file). Set up ElevenLabs or OpenAI in Voice & API to get spoken announcements baked in.');
    }
    if (hasYouTube) warnings.push('• YouTube clips will be silence in the WAV (DRM). They play correctly in live preview.');
    if (hasSpotify) warnings.push('• Spotify tracks will be silence in the WAV (DRM), unless the clip fits inside the first 30 seconds of a track that has a preview available. They play correctly in live preview.');

    if (warnings.length > 0) {
      const lines = [
        'A few things to know about this audio export:',
        '',
        ...warnings,
        '',
        'Continue compiling anyway?',
      ];
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

      // Compute per-source gain factors for loudness normalisation
      const sourceGains = new Map(); // questionId -> gain factor
      if (normaliseLoudness) {
        const sourcesWithRms = questions.filter(q => q.source?.kind === 'file' && q.source.rms > 0);
        if (sourcesWithRms.length >= 2) {
          // Use median RMS as the target — robust against outliers
          const rmsValues = sourcesWithRms.map(q => q.source.rms).sort((a, b) => a - b);
          const targetRms = rmsValues[Math.floor(rmsValues.length / 2)];
          for (const q of sourcesWithRms) {
            // Clamp gain to a reasonable range to prevent over-amplification of near-silent clips
            const gain = Math.min(4, Math.max(0.25, targetRms / q.source.rms));
            sourceGains.set(q.id, gain);
          }
        }
      }

      const ttsItems = timeline.filter(t => t.type === 'tts');
      const ttsBuffers = new Map();
      let ttsBakeMode = 'marker'; // 'marker' | 'eleven' | 'openai'

      if (ttsProvider === 'eleven' && elevenKey) {
        ttsBakeMode = 'eleven';
      } else if (ttsProvider === 'openai' && openaiKey) {
        ttsBakeMode = 'openai';
      }

      if (ttsBakeMode !== 'marker') {
        const providerName = ttsBakeMode === 'eleven' ? 'ElevenLabs' : 'OpenAI';
        let successCount = 0;
        for (let i = 0; i < ttsItems.length; i++) {
          setCompileStatus(`Generating ${providerName} announcement ${i + 1} / ${ttsItems.length}...`);
          setCompileProgress((i / ttsItems.length) * 70);
          try {
            const buf = await renderTTSBuffer(ttsItems[i].text);
            if (buf) {
              ttsBuffers.set(i, buf);
              successCount++;
            } else {
              console.warn(`TTS returned null buffer for: "${ttsItems[i].text.slice(0, 50)}"`);
              ttsBuffers.set(i, await makeMarkerTone(ttsItems[i].duration));
            }
          } catch (err) {
            console.warn(`TTS failed for "${ttsItems[i].text.slice(0, 50)}": ${err.message}`);
            ttsBuffers.set(i, await makeMarkerTone(ttsItems[i].duration));
          }
        }
        if (successCount === 0) {
          alert(`All ${providerName} announcements failed to generate. Check your API key and try again.\n\nThe WAV will be exported with marker tones only.`);
        } else if (successCount < ttsItems.length) {
          console.warn(`${ttsItems.length - successCount} of ${ttsItems.length} announcements fell back to markers.`);
        }
      } else {
        // No premium TTS — markers only
        setCompileStatus('No premium TTS key set — using timing markers for announcements...');
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
      const fadeDur = crossfadeMs / 1000; // seconds
      timeline.forEach((item) => {
        if (item.type === 'audio') {
          const src = offlineCtx.createBufferSource();
          src.buffer = item.buffer;
          // Per-source normalisation gain + fade-in/out envelope
          const gainNode = offlineCtx.createGain();
          const normGain = sourceGains.get(item.questionId) ?? 1;
          // Build envelope: fade in over fadeDur, hold, fade out over fadeDur
          const itemDur = item.duration;
          if (fadeDur > 0 && itemDur > fadeDur * 2.5) {
            gainNode.gain.setValueAtTime(0, item.start);
            gainNode.gain.linearRampToValueAtTime(normGain, item.start + fadeDur);
            gainNode.gain.setValueAtTime(normGain, item.start + itemDur - fadeDur);
            gainNode.gain.linearRampToValueAtTime(0, item.start + itemDur);
          } else {
            gainNode.gain.setValueAtTime(normGain, item.start);
          }
          src.connect(gainNode).connect(offlineCtx.destination);
          if (item.bufferOffset != null) {
            src.start(item.start, item.bufferOffset, item.playDuration);
          } else {
            src.start(item.start);
          }
        } else if (item.type === 'tts') {
          const buf = ttsBuffers.get(ttsIndex);
          ttsIndex++;
          if (buf) {
            const src = offlineCtx.createBufferSource();
            src.buffer = buf;
            // Gentle fade-in/out for TTS too (just 30ms — much shorter than music) to avoid clicks
            const gainNode = offlineCtx.createGain();
            const ttsFade = 0.03;
            const ttsDur = buf.duration;
            if (ttsDur > ttsFade * 2.5) {
              gainNode.gain.setValueAtTime(0, item.start);
              gainNode.gain.linearRampToValueAtTime(1, item.start + ttsFade);
              gainNode.gain.setValueAtTime(1, item.start + ttsDur - ttsFade);
              gainNode.gain.linearRampToValueAtTime(0, item.start + ttsDur);
            }
            src.connect(gainNode).connect(offlineCtx.destination);
            src.start(item.start);
          }
        } else if (item.type === 'spotify' && item.previewUrl && item.endSec <= 30 && spotifyPreviewBuffers.has(item.previewUrl)) {
          const buf = spotifyPreviewBuffers.get(item.previewUrl);
          const src = offlineCtx.createBufferSource();
          src.buffer = buf;
          // Apply fade to spotify clips too
          const gainNode = offlineCtx.createGain();
          if (fadeDur > 0 && item.duration > fadeDur * 2.5) {
            gainNode.gain.setValueAtTime(0, item.start);
            gainNode.gain.linearRampToValueAtTime(1, item.start + fadeDur);
            gainNode.gain.setValueAtTime(1, item.start + item.duration - fadeDur);
            gainNode.gain.linearRampToValueAtTime(0, item.start + item.duration);
          }
          src.connect(gainNode).connect(offlineCtx.destination);
          // Offset within preview clip = item.startSec
          src.start(item.start, item.startSec, item.duration);
        }
      });

      setCompileStatus('Rendering...');
      setCompileProgress(90);
      const rendered = await offlineCtx.startRendering();

      // Encode in the chosen format
      let blob, ext, mimeLabel;
      if (outputFormat === 'wav') {
        setCompileStatus('Encoding WAV...');
        setCompileProgress(97);
        blob = audioBufferToWav(rendered);
        ext = 'wav';
        mimeLabel = 'WAV';
      } else if (outputFormat === 'mp3') {
        setCompileStatus('Encoding MP3 (this takes a while)...');
        setCompileProgress(92);
        try {
          blob = await audioBufferToMp3(rendered, mp3Bitrate, (frac) => {
            setCompileProgress(92 + frac * 7);
          });
          ext = 'mp3';
          mimeLabel = `MP3 ${mp3Bitrate}kbps`;
        } catch (err) {
          alert(`MP3 encoding failed: ${err.message}\n\nFalling back to WAV.`);
          blob = audioBufferToWav(rendered);
          ext = 'wav';
          mimeLabel = 'WAV';
        }
      } else if (outputFormat === 'ogg' || outputFormat === 'webm') {
        setCompileStatus(`Encoding ${outputFormat.toUpperCase()}...`);
        setCompileProgress(94);
        const fmt = pickSupportedCompressedFormat();
        if (!fmt) {
          alert('Your browser doesn\'t support OGG/WebM encoding. Falling back to WAV.');
          blob = audioBufferToWav(rendered);
          ext = 'wav';
          mimeLabel = 'WAV';
        } else {
          try {
            blob = await audioBufferToCompressed(rendered, fmt.mime, 128000);
            ext = fmt.ext;
            mimeLabel = `${ext.toUpperCase()} (Opus)`;
          } catch (err) {
            alert(`${outputFormat.toUpperCase()} encoding failed: ${err.message}\n\nFalling back to WAV.`);
            blob = audioBufferToWav(rendered);
            ext = 'wav';
            mimeLabel = 'WAV';
          }
        }
      } else {
        blob = audioBufferToWav(rendered);
        ext = 'wav';
        mimeLabel = 'WAV';
      }

      const url = URL.createObjectURL(blob);
      if (finalAudioUrl) URL.revokeObjectURL(finalAudioUrl);
      setFinalAudioUrl(url);
      setFinalAudioDuration(rendered.duration);
      setFinalAudioFormat(ext);
      setFinalAudioSize(blob.size);
      setCompileProgress(100);
      setCompileStatus(`Done · ${mimeLabel}`);
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
    a.download = `${examTitle.replace(/[^a-z0-9]+/gi, '_')}.${finalAudioFormat || 'wav'}`;
    a.click();
  };

  // Human-readable file size
  const formatFileSize = (bytes) => {
    if (!bytes) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  };

  const { totalDuration } = buildTimeline();
  const filledCount = questions.filter(q => q.source).length;
  const youtubeCount = questions.filter(q => q.source?.kind === 'youtube').length;
  const spotifyCount = questions.filter(q => q.source?.kind === 'spotify').length;
  const totalMarks = questions.reduce((sum, q) => sum + (q.marks || 0), 0);

  // ===== Keyboard shortcuts ===== (placed here so all referenced state/functions are defined)
  useEffect(() => {
    const handler = (e) => {
      const tag = (e.target.tagName || '').toLowerCase();
      const isEditable = ['input', 'textarea', 'select'].includes(tag) || e.target.isContentEditable;
      if (isEditable) return;
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      switch (e.key) {
        case ' ':
        case 'Spacebar':
          e.preventDefault();
          if (livePlaying) {
            if (livePaused) resumeLive(); else pauseLive();
          } else if (filledCount > 0 && !isCompiling) {
            playLiveFull();
          }
          break;
        case 'n':
        case 'N':
          if (livePlaying) { e.preventDefault(); skipToNextExtract(); }
          break;
        case 'p':
        case 'P':
          if (livePlaying) { e.preventDefault(); skipToPrevExtract(); }
          break;
        case 'k':
        case 'K':
          if (livePlaying) { e.preventDefault(); skipCurrentItem(); }
          break;
        case 'Escape':
          if (livePlaying) { e.preventDefault(); stopAll(); }
          else if (mobileDrawerOpen) { e.preventDefault(); setMobileDrawerOpen(false); }
          else if (showSettings) { e.preventDefault(); setShowSettings(false); }
          break;
        case 't':
        case 'T':
          e.preventDefault();
          toggleTheme();
          break;
        default:
          break;
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [livePlaying, livePaused, filledCount, isCompiling, mobileDrawerOpen, showSettings]);

  const renderSidebarBody = (closeAfter) => {
    const cloudEntries = cloudExams.map(e => ({
      kind: 'cloud',
      id: e.id,
      name: e.name,
      config: e.config,
      shared_with_all: e.shared_with_all,
      savedAt: e.updated_at,
      ownerEmail: e.owner?.email,
      ownerName: e.owner?.display_name,
      isMine: session ? e.owner_id === session.user?.id : false,
    }));
    const localEntries = savedExams.map(e => ({
      kind: 'local',
      id: e.id,
      name: e.name,
      config: e.config,
      savedAt: e.savedAt,
      isMine: true,
    }));
    const allEntries = [...cloudEntries, ...localEntries];

    return (
      <>
        <div className="flex gap-1.5 mt-3">
          <button onClick={() => { newBlankExam(); closeAfter?.(); }}
            className="btn btn-secondary btn-sm flex-1"
            title="Start a new blank exam">
            <Plus size={12} /> New
          </button>
          <button onClick={() => { saveCurrentExam(); closeAfter?.(); }}
            className="btn btn-primary btn-sm flex-1"
            title={myProfile?.approved ? 'Save to cloud workspace' : 'Save to browser'}>
            <Save size={12} /> Save
          </button>
        </div>

        {cloudExamsLoading && (
          <div className="mt-3 flex items-center gap-2 px-1" style={{ fontSize: '11px', color: 'var(--text-dim)' }}>
            <Loader2 size={10} className="animate-spin" /> Syncing workspace…
          </div>
        )}

        <div className="mt-4 space-y-0.5 max-h-[60vh] overflow-y-auto" style={{ marginLeft: '-4px', marginRight: '-4px', paddingLeft: '4px', paddingRight: '4px' }}>
          {allEntries.length === 0 ? (
            <div className="py-6 text-center" style={{ fontSize: '12px', lineHeight: 1.5, color: 'var(--text-dim)' }}>
              <div style={{ marginBottom: '4px' }}>No saved exams yet</div>
              <div style={{ color: 'var(--text-faint)' }}>Click <strong style={{ color: 'var(--text-muted)' }}>Save</strong> above to store this setup.</div>
            </div>
          ) : (
            allEntries.map(entry => (
              <SavedExamRow
                key={`${entry.kind}-${entry.id}`}
                entry={entry}
                onLoad={() => { loadSavedExam(entry); closeAfter?.(); }}
                onUpdate={entry.isMine ? () => updateExistingExam(entry) : null}
                onRename={entry.isMine ? () => renameSavedExam(entry) : null}
                onDelete={entry.isMine ? () => deleteSavedExam(entry) : null}
                onToggleShare={entry.kind === 'cloud' && entry.isMine ? () => toggleShareWithAll(entry) : null}
              />
            ))
          )}
        </div>

        <div className="mt-4 pt-4 px-1" style={{ borderTop: '0.5px solid var(--border)', fontSize: '11px', lineHeight: 1.5, color: 'var(--text-dim)' }}>
          {myProfile?.approved
            ? 'Cloud-saved exams are accessible from any device. Local exams (LOCAL badge) stay in this browser.'
            : 'Exams are saved in your browser only until your account is approved.'}
        </div>
      </>
    );
  };

  // ===== Auth gating: show auth screen / waiting-for-approval / app =====
  // Style content must render in every branch, since the auth/pending screens render before
  // the main app and otherwise wouldn't see the CSS variables and utility classes.
  const globalStyles = (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Geist:wght@400;500;600;700&family=Geist+Mono:wght@400;500&display=swap');
      body { margin: 0; background: var(--bg-base); }

      :root, [data-theme="dark"] {
        /* Base — softer charcoal, not pure black. Reads as premium, not gaming. */
        --bg-base: #0e0e13;
        --bg-gradient: radial-gradient(ellipse 1200px 800px at top left, #1a1a24 0%, #0e0e13 60%);
        --bg-image-auth: url('/bg/auth-dark.webp');
        --bg-image-main: url('/bg/main-dark.webp');
        --bg-overlay: linear-gradient(180deg, rgba(14,14,19,0.62) 0%, rgba(14,14,19,0.78) 100%);

        /* Surfaces — layered, all slightly translucent so content feels lifted */
        --surface: rgba(255,255,255,0.025);
        --surface-2: rgba(255,255,255,0.015);
        --surface-elev: rgba(255,255,255,0.05);
        --surface-solid: #1a1a22;
        --surface-solid-2: #15151c;

        /* Borders — quieter than before, two levels only */
        --border: rgba(255,255,255,0.06);
        --border-strong: rgba(255,255,255,0.12);

        /* Text — refined four-level hierarchy */
        --text: #ececf1;
        --text-muted: #9a9aa6;
        --text-dim: #6b6b75;
        --text-faint: #4a4a52;

        /* Accent — purple/indigo. Used sparingly. */
        --accent: #7c7df0;
        --accent-soft: #9ea0f5;
        --accent-strong: #6366f1;
        --accent-bg-on: #ffffff;
        --accent-tint: rgba(124,125,240,0.08);
        --accent-tint-strong: rgba(124,125,240,0.14);
        --accent-border: rgba(124,125,240,0.22);
        --accent-glow: rgba(124,125,240,0.18);
        --accent2: #22d3ee;
        --accent2-tint: rgba(34,211,238,0.12);

        /* Semantic colours */
        --success: #4ade80;
        --success-soft: #86efac;
        --success-tint: rgba(74,222,128,0.08);
        --warning: #fbbf24;
        --warning-tint: rgba(251,191,36,0.08);
        --danger: #f87171;
        --danger-tint: rgba(248,113,113,0.08);

        --header-bg: rgba(14,14,19,0.7);
        --input-bg: rgba(255,255,255,0.025);
        --input-bg-focus: rgba(255,255,255,0.04);

        /* Elevations */
        --shadow-sm: 0 1px 2px rgba(0,0,0,0.3);
        --shadow-card: 0 1px 0 rgba(255,255,255,0.04) inset, 0 1px 3px rgba(0,0,0,0.3), 0 8px 32px rgba(0,0,0,0.35);
        --shadow-lg: 0 4px 8px rgba(0,0,0,0.3), 0 24px 48px rgba(0,0,0,0.45);
        --logo-shadow: 0 1px 0 rgba(255,255,255,0.15) inset, 0 4px 12px rgba(99,102,241,0.3);

        /* Spacing scale — multiples of 4 */
        --r-sm: 6px;
        --r-md: 8px;
        --r-lg: 12px;
        --r-xl: 16px;

        --ink: #2a2520;
        --waveform-bg: rgba(0,0,0,0.25);
        --code-bg: #0a0a0e;
        --code-text: #d5d5dc;
        --scrollbar: rgba(255,255,255,0.08);
        --scrollbar-hover: rgba(255,255,255,0.15);
      }

      [data-theme="light"] {
        --bg-base: #f0eef0;
        --bg-gradient: radial-gradient(ellipse 1400px 900px at top left, #e8e4f5 0%, #f0eef0 50%, #ebe9eb 100%);
        --bg-image-auth: url('/bg/auth-light.webp');
        --bg-image-main: url('/bg/main-light.webp');
        --bg-overlay: linear-gradient(180deg, rgba(240,238,240,0.55) 0%, rgba(240,238,240,0.7) 100%);

        --surface: #ffffff;
        --surface-2: #fafaf9;
        --surface-elev: rgba(0,0,0,0.035);
        --surface-solid: #ffffff;
        --surface-solid-2: #fafaf9;

        --border: rgba(0,0,0,0.08);
        --border-strong: rgba(0,0,0,0.16);

        --text: #1a1a1f;
        --text-muted: #58585e;
        --text-dim: #8e8e94;
        --text-faint: #b0b0b5;

        --accent: #5b5cd6;
        --accent-soft: #6366f1;
        --accent-strong: #4f46e5;
        --accent-bg-on: #ffffff;
        --accent-tint: rgba(91,92,214,0.06);
        --accent-tint-strong: rgba(91,92,214,0.12);
        --accent-border: rgba(91,92,214,0.28);
        --accent-glow: rgba(91,92,214,0.15);
        --accent2: #0891b2;
        --accent2-tint: rgba(8,145,178,0.08);

        --success: #16a34a;
        --success-soft: #22c55e;
        --success-tint: rgba(22,163,74,0.06);
        --warning: #d97706;
        --warning-tint: rgba(217,119,6,0.06);
        --danger: #dc2626;
        --danger-tint: rgba(220,38,38,0.06);

        --header-bg: rgba(255,255,255,0.85);
        --input-bg: #ffffff;
        --input-bg-focus: #ffffff;

        --shadow-sm: 0 1px 2px rgba(0,0,0,0.04);
        --shadow-card: 0 1px 2px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.05), 0 16px 40px rgba(0,0,0,0.06), 0 0 0 0.5px rgba(0,0,0,0.06);
        --shadow-lg: 0 4px 12px rgba(0,0,0,0.08), 0 24px 48px rgba(0,0,0,0.1);
        --logo-shadow: 0 1px 0 rgba(255,255,255,0.4) inset, 0 4px 12px rgba(79,70,229,0.25);

        --r-sm: 6px;
        --r-md: 8px;
        --r-lg: 12px;
        --r-xl: 16px;

        --ink: #f0f0f3;
        --waveform-bg: rgba(0,0,0,0.03);
        --code-bg: #1a1a1f;
        --code-text: #f0f0f3;
        --scrollbar: rgba(0,0,0,0.1);
        --scrollbar-hover: rgba(0,0,0,0.18);
      }

      .display-font { font-family: 'Geist', system-ui, sans-serif; letter-spacing: -0.015em; }
      .mono-font { font-family: 'Geist Mono', 'Menlo', monospace; }

      /* Background-image layers for auth and main screens.
         Image sits at the bottom with a translucent overlay on top to keep UI legible.
         The image stays fixed while page content scrolls. */
      .bg-image-auth {
        background:
          var(--bg-overlay),
          var(--bg-image-auth) center/cover no-repeat fixed,
          var(--bg-base);
      }
      .bg-image-main {
        background:
          var(--bg-overlay),
          var(--bg-image-main) center/cover no-repeat fixed,
          var(--bg-base);
      }

      .accent { color: var(--accent-soft); }
      .accent-bg { background: var(--accent-strong); color: var(--accent-bg-on); box-shadow: 0 1px 0 rgba(255,255,255,0.15) inset; }
      .accent2 { color: var(--accent2); }

      /* Button hierarchy ------------------------------------------------ */
      .btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        font-family: 'Geist', system-ui, sans-serif;
        font-size: 13px;
        font-weight: 500;
        line-height: 1;
        padding: 9px 14px;
        border-radius: var(--r-md);
        border: 0.5px solid transparent;
        cursor: pointer;
        transition: background 0.15s, border-color 0.15s, color 0.15s, transform 0.05s;
        white-space: nowrap;
      }
      .btn:active:not(:disabled) { transform: translateY(0.5px); }
      .btn:focus-visible {
        outline: none;
        box-shadow: 0 0 0 2px var(--bg-base), 0 0 0 4px var(--accent-glow);
      }
      .btn:disabled { opacity: 0.4; cursor: not-allowed; }

      .btn-primary {
        background: var(--accent-strong);
        color: var(--accent-bg-on);
        font-weight: 600;
        box-shadow: 0 1px 0 rgba(255,255,255,0.18) inset, 0 1px 2px rgba(0,0,0,0.1);
      }
      .btn-primary:hover:not(:disabled) {
        background: var(--accent);
        box-shadow: 0 1px 0 rgba(255,255,255,0.22) inset, 0 2px 6px rgba(99,102,241,0.3);
      }

      .btn-secondary {
        background: var(--surface-elev);
        color: var(--text);
        border-color: var(--border);
      }
      .btn-secondary:hover:not(:disabled) {
        background: var(--surface-elev);
        border-color: var(--border-strong);
      }

      .btn-ghost {
        background: transparent;
        color: var(--text-muted);
      }
      .btn-ghost:hover:not(:disabled) {
        background: var(--surface-elev);
        color: var(--text);
      }

      .btn-danger-ghost {
        background: transparent;
        color: var(--text-muted);
      }
      .btn-danger-ghost:hover:not(:disabled) {
        background: var(--danger-tint);
        color: var(--danger);
      }

      .btn-icon { padding: 8px; gap: 0; }
      .btn-sm { padding: 6px 10px; font-size: 12px; }
      .btn-icon.btn-sm { padding: 6px; }

      /* Surfaces --------------------------------------------------------- */
      .paper { background: var(--surface); }
      .paper-elev { background: var(--surface-2); }
      .hairline { border: 0.5px solid var(--border); }
      .ink-shadow { box-shadow: var(--shadow-card); }

      /* Card — standard container */
      .card {
        background: var(--surface);
        border: 0.5px solid var(--border);
        border-radius: var(--r-lg);
      }
      .card-elev {
        background: var(--surface);
        border: 0.5px solid var(--border);
        border-radius: var(--r-lg);
        box-shadow: var(--shadow-card);
      }

      /* Pills & badges --------------------------------------------------- */
      .pill {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        font-size: 11px;
        font-weight: 500;
        padding: 3px 8px;
        border-radius: 999px;
        background: var(--surface-elev);
        color: var(--text-muted);
        border: 0.5px solid var(--border);
      }
      .pill-accent { background: var(--accent-tint); color: var(--accent-soft); border-color: var(--accent-border); }
      .pill-success { background: var(--success-tint); color: var(--success); border-color: rgba(74,222,128,0.2); }
      .pill-warning { background: var(--warning-tint); color: var(--warning); border-color: rgba(251,191,36,0.2); }
      .pill-muted { background: transparent; color: var(--text-dim); border-color: var(--border); }

      /* Section label — replaces uppercase tracking-wider noise */
      .section-label {
        font-family: 'Geist', system-ui, sans-serif;
        font-size: 11px;
        font-weight: 500;
        color: var(--text-dim);
        letter-spacing: 0.02em;
      }
      .field-label {
        font-family: 'Geist', system-ui, sans-serif;
        font-size: 12px;
        font-weight: 500;
        color: var(--text-muted);
        display: block;
        margin-bottom: 6px;
      }

      input[type="number"], input[type="text"], input[type="password"], input[type="email"], textarea, select {
        background: var(--input-bg);
        border: 0.5px solid var(--border);
        padding: 8px 12px;
        font-family: inherit;
        color: var(--text);
        border-radius: var(--r-md);
        transition: border-color 0.15s, background 0.15s, box-shadow 0.15s;
        font-size: 14px;
      }
      input:hover:not(:disabled):not(:focus), textarea:hover:not(:disabled):not(:focus), select:hover:not(:disabled):not(:focus) {
        border-color: var(--border-strong);
      }
      input:focus, textarea:focus, select:focus {
        outline: none;
        background: var(--input-bg-focus);
        border-color: var(--accent-border);
        box-shadow: 0 0 0 3px var(--accent-glow);
      }
      input::placeholder, textarea::placeholder { color: var(--text-faint); }

      button:focus-visible:not(.btn) {
        outline: none;
        box-shadow: 0 0 0 2px var(--bg-base), 0 0 0 4px var(--accent-glow);
      }
      button { transition: all 0.15s; cursor: pointer; }
      button:disabled { opacity: 0.4; cursor: not-allowed; }

      .question-card {
        transition: border-color 0.15s, box-shadow 0.15s;
        background: var(--surface);
        border: 0.5px solid var(--border);
        border-radius: var(--r-lg);
        position: relative;
      }
      .question-card:hover {
        border-color: var(--border-strong);
        box-shadow: var(--shadow-sm);
      }
      .question-card.missing-audio::before {
        content: '';
        position: absolute;
        left: 0;
        top: 16px;
        bottom: 16px;
        width: 2px;
        background: var(--warning);
        border-radius: 0 2px 2px 0;
        opacity: 0.7;
      }
      .question-card.has-audio::before {
        content: '';
        position: absolute;
        left: 0;
        top: 16px;
        bottom: 16px;
        width: 2px;
        background: var(--accent-strong);
        border-radius: 0 2px 2px 0;
        opacity: 0.5;
      }

      .drop-zone {
        background: var(--surface-2);
        border: 1px dashed var(--border-strong);
        border-radius: var(--r-md);
        transition: all 0.15s;
      }
      .drop-zone:hover {
        background: var(--surface-elev);
        border-color: var(--accent-border);
      }
      .drop-zone.has-source {
        background: var(--accent-tint);
        border: 0.5px solid var(--accent-border);
        border-style: solid;
      }

      @keyframes shimmer { 0% { background-position: -200% 0; } 100% { background-position: 200% 0; } }
      .progress-bar {
        background: linear-gradient(90deg, var(--accent) 0%, var(--accent-soft) 50%, var(--accent) 100%);
        background-size: 200% 100%;
        animation: shimmer 2s linear infinite;
      }

      .tab {
        padding: 6px 12px;
        border: 0.5px solid var(--border);
        background: var(--surface-elev);
        color: var(--text-muted);
        border-radius: var(--r-md);
        font-size: 12px;
        font-weight: 500;
        transition: all 0.15s;
        cursor: pointer;
      }
      .tab:hover:not(.active):not(:disabled) {
        background: var(--surface-elev);
        color: var(--text);
        border-color: var(--border-strong);
      }
      .tab.active {
        background: var(--accent-strong);
        color: var(--accent-bg-on);
        border-color: var(--accent-strong);
        box-shadow: 0 1px 0 rgba(255,255,255,0.15) inset;
      }

      .hover-glow:hover { background: var(--surface-elev) !important; }

      .yt-hidden { position: fixed; bottom: -200px; right: 10px; opacity: 0.01; pointer-events: none; }

      details > summary { list-style: none; cursor: pointer; }
      details > summary::-webkit-details-marker { display: none; }

      input[type="range"] { background: transparent; padding: 0; border: none; }
      input[type="range"]::-webkit-slider-runnable-track { background: var(--surface-elev); height: 4px; border-radius: 2px; }
      input[type="range"]::-webkit-slider-thumb { appearance: none; width: 14px; height: 14px; background: var(--accent-soft); border-radius: 50%; margin-top: -5px; cursor: pointer; }
      input[type="range"]::-moz-range-track { background: var(--surface-elev); height: 4px; border-radius: 2px; }
      input[type="range"]::-moz-range-thumb { width: 14px; height: 14px; background: var(--accent-soft); border-radius: 50%; border: none; cursor: pointer; }

      ::-webkit-scrollbar { width: 8px; height: 8px; }
      ::-webkit-scrollbar-track { background: transparent; }
      ::-webkit-scrollbar-thumb { background: var(--scrollbar); border-radius: 4px; }
      ::-webkit-scrollbar-thumb:hover { background: var(--scrollbar-hover); }

      @keyframes slideIn { from { transform: translateX(-100%); } to { transform: translateX(0); } }
      @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
      .drawer-overlay {
        position: fixed; inset: 0; background: rgba(0,0,0,0.5);
        z-index: 50; animation: fadeIn 0.2s ease;
      }
      .drawer-panel {
        position: fixed; left: 0; top: 0; bottom: 0;
        width: 280px; max-width: 85vw;
        background: var(--bg-base);
        border-right: 0.5px solid var(--border);
        z-index: 51; animation: slideIn 0.2s ease;
        overflow-y: auto;
      }

      @media (max-width: 640px) {
        .responsive-grid-2 { grid-template-columns: 1fr !important; }
        .responsive-grid-3 { grid-template-columns: 1fr 1fr !important; }
        .responsive-grid-4 { grid-template-columns: 1fr 1fr !important; }
        .hide-mobile { display: none !important; }
        .stack-mobile { flex-direction: column !important; align-items: stretch !important; }
        .stack-mobile > * { width: 100%; }
      }
      @media (min-width: 641px) {
        .show-mobile-only { display: none !important; }
      }
    `}</style>
  );

  if (authLoading) {
    return (
      <>
        {globalStyles}
        <div className="min-h-screen flex items-center justify-center" style={{ background: 'var(--bg-gradient)', color: 'var(--text)' }}>
          <Loader2 size={20} className="animate-spin" />
        </div>
      </>
    );
  }
  if (!session) {
    return (
      <>
        {globalStyles}
        <AuthScreen theme={theme} toggleTheme={toggleTheme} />
      </>
    );
  }
  if (myProfile && !myProfile.approved) {
    return (
      <>
        {globalStyles}
        <PendingApprovalScreen profile={myProfile} onSignOut={signOut} theme={theme} toggleTheme={toggleTheme} />
      </>
    );
  }

  return (
    <div className="min-h-screen bg-image-main" style={{
      fontFamily: "'Geist', system-ui, -apple-system, sans-serif",
      color: 'var(--text)',
    }}>
      {globalStyles}

      <div className="yt-hidden"><div ref={ytContainerRef}></div></div>

      {/* Toast notifications */}
      {showAdminPanel && (
        <AdminPanel onClose={() => setShowAdminPanel(false)} onChange={reloadCloudExams} />
      )}

      {showHelpPanel && (
        <HelpPanel onClose={() => setShowHelpPanel(false)} />
      )}

      {showWelcomeBanner && (
        <WelcomeBanner onDismiss={dismissWelcome} onOpenHelp={() => { dismissWelcome(); setShowHelpPanel(true); }} />
      )}

      {shareToastMessage && (
        <div style={{
          position: 'fixed', bottom: '24px', left: '50%', transform: 'translateX(-50%)',
          background: 'var(--surface)', color: 'var(--text)',
          border: '0.5px solid var(--accent-border)',
          padding: '12px 18px', borderRadius: '10px',
          boxShadow: 'var(--shadow-card)',
          zIndex: 100, maxWidth: '90vw',
          display: 'flex', alignItems: 'center', gap: '10px',
          fontSize: '13px',
        }}>
          <Check size={14} className="accent" />
          <span>{shareToastMessage}</span>
        </div>
      )}

      <header className="hairline" style={{ borderBottom: '0.5px solid var(--border)', background: 'var(--header-bg)', backdropFilter: 'blur(8px)', position: 'sticky', top: 0, zIndex: 20 }}>
        <div className="max-w-7xl mx-auto px-4 sm:px-8 py-4 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3 min-w-0">
            <button
              onClick={() => setMobileDrawerOpen(true)}
              className="show-mobile-only btn btn-ghost btn-icon"
              aria-label="Open sidebar">
              <Menu size={18} />
            </button>
            <div style={{
              width: '32px', height: '32px',
              background: 'linear-gradient(135deg, var(--accent-strong) 0%, #4338ca 100%)',
              borderRadius: 'var(--r-md)',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              boxShadow: 'var(--logo-shadow)',
              flexShrink: 0,
            }}>
              <Music size={16} color="#ffffff" strokeWidth={2.2} />
            </div>
            <div className="min-w-0">
              <h1 className="display-font font-semibold leading-tight truncate" style={{ fontSize: '15px', letterSpacing: '-0.01em' }}>Aural Composer</h1>
              <div className="hide-mobile" style={{ fontSize: '11px', color: 'var(--text-dim)', marginTop: '2px' }}>Listening exam composer</div>
            </div>
          </div>
          <div className="flex items-center gap-1 flex-shrink-0">
            <button onClick={() => setShowHelpPanel(true)}
              className="btn btn-ghost btn-icon"
              title="Help & how-to guide"
              aria-label="Help">
              <HelpCircle size={15} />
            </button>
            {myProfile?.is_admin && (
              <button onClick={() => setShowAdminPanel(true)}
                className="btn btn-ghost btn-icon"
                title="Manage users (admin only)"
                aria-label="Admin panel">
                <Users size={15} />
              </button>
            )}
            <button onClick={toggleTheme}
              className="btn btn-ghost btn-icon"
              title={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
              aria-label="Toggle theme">
              {theme === 'dark' ? <Sun size={15} /> : <Moon size={15} />}
            </button>
            <button onClick={() => setShowSettings(!showSettings)}
              className={showSettings ? "btn btn-secondary" : "btn btn-ghost"}
              style={showSettings ? { background: 'var(--accent-tint-strong)', color: 'var(--accent-soft)', borderColor: 'var(--accent-border)' } : undefined}>
              <Settings size={14} />
              <span className="hide-mobile">Voice & API</span>
            </button>
            <button onClick={signOut}
              className="btn btn-ghost btn-icon"
              title={`Signed in as ${myProfile?.email || session?.user?.email || ''} — click to sign out`}
              aria-label="Sign out">
              <LogOut size={15} />
            </button>
          </div>
        </div>
      </header>

      {showSettings && (
        <div style={{ background: 'var(--surface-2)', borderBottom: '0.5px solid var(--border)' }}>
          <div className="max-w-7xl mx-auto px-4 sm:px-8 py-4 sm:py-6">
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
            <div className="mt-6 pt-6" style={{ borderTop: '1px dashed rgba(255,255,255,0.1)' }}>
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
                      style={{ color: 'var(--surface)', borderRadius: '2px' }}>
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
                  <code className="mono-font block mt-1 p-2" style={{ background: 'var(--text)', color: 'var(--surface)', borderRadius: '2px', wordBreak: 'break-all' }}>
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

      <div className="flex max-w-7xl mx-auto" style={{ minHeight: 'calc(100vh - 90px)' }}>
        {/* Sidebar (desktop) */}
        <aside className="hide-mobile" style={{
          width: sidebarOpen ? '256px' : '48px',
          flexShrink: 0,
          borderRight: '0.5px solid var(--border)',
          transition: 'width 0.2s ease',
        }}>
          <div className="sticky top-0" style={{ padding: '16px 12px' }}>
            <button onClick={() => setSidebarOpen(!sidebarOpen)}
              className="w-full btn btn-ghost btn-sm"
              style={{ justifyContent: sidebarOpen ? 'space-between' : 'center' }}
              title={sidebarOpen ? 'Hide sidebar' : 'Show sidebar'}>
              <span className="flex items-center gap-2">
                <ListMusic size={13} />
                {sidebarOpen && <span style={{ fontWeight: 500 }}>Saved exams</span>}
              </span>
              {sidebarOpen && <ChevronRight size={12} style={{ transform: 'rotate(180deg)', color: 'var(--text-dim)' }} />}
            </button>
            {sidebarOpen && renderSidebarBody(null)}
          </div>
        </aside>

        {/* Mobile drawer */}
        {mobileDrawerOpen && (
          <>
            <div className="drawer-overlay" onClick={() => setMobileDrawerOpen(false)} />
            <div className="drawer-panel">
              <div style={{ padding: '16px 12px' }}>
                <div className="flex items-center justify-between mb-3">
                  <div className="flex items-center gap-2" style={{ fontSize: '13px', fontWeight: 500, color: 'var(--text)' }}>
                    <ListMusic size={14} /> Saved exams
                  </div>
                  <button onClick={() => setMobileDrawerOpen(false)}
                    className="btn btn-ghost btn-icon btn-sm"
                    aria-label="Close drawer">
                    <X size={14} />
                  </button>
                </div>
                {renderSidebarBody(() => setMobileDrawerOpen(false))}
              </div>
            </div>
          </>
        )}

        {/* Main content */}
        <main className="flex-1 min-w-0 px-4 sm:px-8 py-6 sm:py-8">

        {/* Quick start strip - dismissible */}
        {showQuickStart && (
          <div className="mb-6 card flex items-center gap-3" style={{
            padding: '12px 16px',
            background: 'var(--surface-2)',
          }}>
            <div style={{
              width: '28px', height: '28px',
              background: 'var(--accent-tint-strong)',
              border: '0.5px solid var(--accent-border)',
              borderRadius: 'var(--r-md)',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              flexShrink: 0,
            }}>
              <Sparkles size={13} className="accent" />
            </div>
            <div className="flex-1" style={{ fontSize: '13px', color: 'var(--text-muted)', lineHeight: 1.5 }}>
              Drop an exam PDF below, add audio to each extract, preview, then compile.{' '}
              <button onClick={() => setShowHelpPanel(true)}
                style={{ background: 'transparent', border: 'none', color: 'var(--accent-soft)', padding: 0, font: 'inherit', cursor: 'pointer', textDecoration: 'underline' }}>
                Full guide
              </button>
            </div>
            <button onClick={dismissQuickStart}
              className="btn btn-ghost btn-icon btn-sm"
              title="Dismiss"
              aria-label="Dismiss quick start tip">
              <X size={12} />
            </button>
          </div>
        )}

        {/* PDF + Save/Load toolbar */}
        <section className="mb-8 card-elev" style={{ padding: '20px' }}>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            {/* PDF upload */}
            <div>
              <div className="section-label mb-2">Start from a paper</div>
              <PdfDropZone onFile={handleExamPdfDrop} parsing={pdfParsing} disabled={isCompiling || livePlaying} />
              {pdfDetectionInfo && (
                <div className="mt-2 flex items-center gap-1.5" style={{ fontSize: '12px', color: 'var(--success)' }}>
                  <Check size={12} /> Loaded {pdfDetectionInfo.count} extracts
                  {pdfDetectionInfo.marksFound && ' with marks'}
                </div>
              )}
            </div>

            {/* Save / Load */}
            <div>
              <div className="section-label mb-2">Or resume a saved exam</div>
              <div className="flex gap-1.5 flex-wrap">
                <button onClick={saveExamConfig} className="btn btn-secondary btn-sm">
                  <Save size={12} /> Save config
                </button>
                <label className="btn btn-secondary btn-sm" style={{ cursor: 'pointer' }}>
                  <FolderOpen size={12} /> Load config
                  <input type="file" accept=".json,application/json" className="hidden"
                    onChange={e => { loadExamConfig(e.target.files[0]); e.target.value = ''; }} />
                </label>
                <button onClick={shareAsUrl} className="btn btn-secondary btn-sm"
                  title="Copy a share link that loads this exam (audio files not included)">
                  <Link2 size={12} /> Share link
                </button>
              </div>
              <div className="mt-2" style={{ fontSize: '12px', color: 'var(--text-dim)', lineHeight: 1.5 }}>
                Saves all settings except uploaded audio (which need to be re-uploaded). YouTube and Spotify clips are saved with their URLs and timestamps.
              </div>
            </div>
          </div>
        </section>

        <section className="mb-10">
          <div className="section-label mb-3">Exam</div>
          <input type="text" value={examTitle} onChange={e => setExamTitle(e.target.value)}
            placeholder="Untitled exam"
            className="display-font w-full bg-transparent border-none p-0"
            style={{
              fontSize: '28px',
              fontWeight: 600,
              letterSpacing: '-0.02em',
              color: 'var(--text)',
              paddingBottom: '8px',
              borderBottom: '0.5px solid var(--border)',
            }} />

          {/* Compact stats dashboard */}
          <div className="card mt-5" style={{ padding: '14px 18px' }}>
            <div className="flex flex-wrap items-center" style={{ gap: '4px 28px' }}>
              <StatCell label="Extracts" value={questions.length} />
              <StatCell label="Audio loaded" value={`${filledCount} of ${questions.length}`} accent={filledCount === questions.length && questions.length > 0} />
              {totalMarks > 0 && <StatCell label="Marks" value={totalMarks} icon={<Award size={11} />} />}
              <StatCell label="Runtime" value={formatTime(totalDuration)} />
              {youtubeCount > 0 && <StatCell label="YouTube" value={youtubeCount} icon={<Youtube size={11} />} />}
              {spotifyCount > 0 && <StatCell label="Spotify" value={spotifyCount} icon={<Music size={11} />} />}
              <div className="flex items-center gap-2 ml-auto" style={{ paddingLeft: '12px', borderLeft: '0.5px solid var(--border)' }}>
                <label className="section-label" style={{ marginBottom: 0 }}>Reading time</label>
                <input type="number" value={readingTime} onChange={e => setReadingTime(Math.max(0, parseInt(e.target.value) || 0))}
                  className="text-sm" style={{ width: '64px', padding: '4px 8px', fontSize: '13px' }} />
                <span style={{ fontSize: '12px', color: 'var(--text-dim)' }}>sec</span>
              </div>
            </div>
          </div>
        </section>

        <section className="mb-8 card-elev">
          <button onClick={() => setShowScript(!showScript)}
            className="w-full flex items-center justify-between"
            style={{
              background: 'transparent',
              borderBottom: showScript ? '0.5px solid var(--border)' : 'none',
              padding: '16px 20px',
              border: 'none',
              borderRadius: showScript ? 0 : 'var(--r-lg)',
              cursor: 'pointer',
            }}>
            <div className="text-left">
              <div className="section-label mb-1">Optional</div>
              <h2 className="display-font font-semibold" style={{ fontSize: '17px', letterSpacing: '-0.01em' }}>Announcement script</h2>
            </div>
            <div className="flex items-center gap-1.5" style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
              <span>{showScript ? 'Hide' : 'Customise'}</span>
              <ChevronDown size={14} style={{ transform: showScript ? 'rotate(180deg)' : 'none', transition: 'transform 0.15s' }} />
            </div>
          </button>

          {showScript && (
            <div className="p-5 pt-2 space-y-4">
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Opening announcement (played at the very start, before reading time)</label>
                <textarea value={script.opening} onChange={e => setScript({ ...script, opening: e.target.value })} rows={3}
                  className="w-full text-sm" style={{ resize: 'vertical' }} />
              </div>
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">After reading time (just before the first extract)</label>
                <textarea value={script.postReading} onChange={e => setScript({ ...script, postReading: e.target.value })} rows={2}
                  className="w-full text-sm" style={{ resize: 'vertical' }} />
              </div>
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">
                  Between plays template — placeholders: <code className="mono-font" style={{ background: 'var(--accent-tint)', padding: '1px 4px', borderRadius: '2px' }}>{'{ord}'}</code> (second/third…) · <code className="mono-font" style={{ background: 'var(--accent-tint)', padding: '1px 4px', borderRadius: '2px' }}>{'{n}'}</code> (2/3…) · <code className="mono-font" style={{ background: 'var(--accent-tint)', padding: '1px 4px', borderRadius: '2px' }}>{'{final}'}</code> (auto-adds " and final" on the last play)
                </label>
                <input type="text" value={script.betweenPlays} onChange={e => setScript({ ...script, betweenPlays: e.target.value })}
                  className="w-full text-sm" />
                <div className="text-xs opacity-60 mt-1">
                  Example for a 3-play extract → 2nd: <em>"{renderBetweenPlays(script.betweenPlays, 2, false)}"</em> · 3rd: <em>"{renderBetweenPlays(script.betweenPlays, 3, true)}"</em>
                </div>
              </div>
              <div>
                <label className="mono-font text-xs uppercase tracking-wider opacity-50 block mb-1">Closing announcement (after the final extract)</label>
                <textarea value={script.ending} onChange={e => setScript({ ...script, ending: e.target.value })} rows={2}
                  className="w-full text-sm" style={{ resize: 'vertical' }} />
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
                <h2 className="display-font text-lg font-semibold flex items-center gap-2">
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
                style={{ color: 'var(--surface)', borderRadius: '2px' }}>
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
              <div className="mt-4 space-y-2 max-h-96 overflow-y-auto pr-2" style={{ borderTop: '1px solid rgba(255,255,255,0.07)', paddingTop: '12px' }}>
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
            <h2 className="display-font font-semibold" style={{ fontSize: '17px', letterSpacing: '-0.01em' }}>Extracts</h2>
            <div style={{ fontSize: '12px', color: 'var(--text-dim)' }}>File · YouTube · Spotify</div>
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
              className="w-full btn btn-ghost"
              style={{
                padding: '14px',
                border: '0.5px dashed var(--border-strong)',
                borderRadius: 'var(--r-lg)',
                color: 'var(--text-muted)',
              }}>
              <Plus size={14} /> Add extract
            </button>
          </div>
        </section>

        <section className="mt-12 paper ink-shadow" style={{ borderRadius: '4px', padding: '32px' }}>
          <div className="flex items-baseline justify-between mb-6 flex-wrap gap-4">
            <div>
              <div className="mono-font text-xs uppercase tracking-widest opacity-60 mb-2">Output</div>
              <h2 className="display-font text-2xl font-semibold">Compile & Export</h2>
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
              <div className="h-1 hairline rounded overflow-hidden" style={{ background: 'var(--surface-elev)' }}>
                <div className="progress-bar h-full transition-all duration-300" style={{ width: `${compileProgress}%` }} />
              </div>
            </div>
          )}

          <div className="flex flex-wrap gap-2">
            {!livePlaying ? (
              <button onClick={playLiveFull} disabled={isCompiling || filledCount === 0}
                className="btn btn-secondary">
                <Play size={14} /> Preview full exam (live)
              </button>
            ) : (
              <div className="flex items-stretch" style={{
                border: '0.5px solid var(--border)',
                borderRadius: 'var(--r-md)',
                overflow: 'hidden',
              }}>
                <button onClick={skipToPrevExtract}
                  className="flex items-center justify-center px-3"
                  style={{ background: 'var(--surface-elev)', borderRight: '0.5px solid var(--border)', color: 'var(--text)', border: 'none', cursor: 'pointer' }}
                  title="Jump to previous extract">
                  <SkipBack size={14} />
                </button>
                <button onClick={livePaused ? resumeLive : pauseLive}
                  className="flex items-center gap-1.5 px-4"
                  style={{
                    background: livePaused ? 'var(--accent-strong)' : 'var(--surface-elev)',
                    color: livePaused ? 'var(--accent-bg-on)' : 'var(--text)',
                    borderRight: '0.5px solid var(--border)',
                    fontSize: '13px',
                    fontWeight: 500,
                    border: 'none',
                    cursor: 'pointer',
                  }}
                  title={livePaused ? 'Resume' : 'Pause'}>
                  {livePaused ? <Play size={14} /> : <Pause size={14} />}
                  {livePaused ? 'Resume' : 'Pause'}
                </button>
                <button onClick={skipCurrentItem}
                  className="flex items-center justify-center px-3"
                  style={{ background: 'var(--surface-elev)', borderRight: '0.5px solid var(--border)', color: 'var(--text)', border: 'none', cursor: 'pointer' }}
                  title="Skip current segment (announcement, silence, or audio)">
                  <ChevronsRight size={14} />
                </button>
                <button onClick={skipToNextExtract}
                  className="flex items-center justify-center px-3"
                  style={{ background: 'var(--surface-elev)', borderRight: '0.5px solid var(--border)', color: 'var(--text)', border: 'none', cursor: 'pointer' }}
                  title="Jump to next extract">
                  <SkipForward size={14} />
                </button>
                <button onClick={playLiveFull}
                  className="flex items-center gap-1.5 px-4"
                  style={{ background: 'var(--danger-tint)', color: 'var(--danger)', fontSize: '12px', fontWeight: 500, border: 'none', cursor: 'pointer' }}
                  title="Stop preview">
                  <Square size={12} /> Stop
                </button>
              </div>
            )}

            <div className="flex items-stretch" style={{
              border: '0.5px solid var(--border)',
              borderRadius: 'var(--r-md)',
              overflow: 'hidden',
            }}>
              <select value={outputFormat} onChange={e => setOutputFormat(e.target.value)}
                disabled={isCompiling || filledCount === 0 || livePlaying}
                style={{
                  padding: '8px 28px 8px 12px',
                  border: 'none',
                  background: 'var(--surface-elev)',
                  color: 'var(--text)',
                  borderRight: '0.5px solid var(--border)',
                  borderRadius: 0,
                  appearance: 'none',
                  WebkitAppearance: 'none',
                  fontSize: '13px',
                  fontWeight: 500,
                  backgroundImage: `url("data:image/svg+xml;charset=utf8,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%239a9aa6' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E")`,
                  backgroundRepeat: 'no-repeat',
                  backgroundPosition: 'right 10px center',
                }}>
                <option value="mp3">MP3</option>
                <option value="wav">WAV</option>
                <option value="ogg">OGG</option>
              </select>
              <button onClick={compileAudio} disabled={isCompiling || filledCount === 0 || livePlaying}
                className="flex items-center gap-1.5 px-5"
                style={{
                  background: 'var(--accent-strong)',
                  color: 'var(--accent-bg-on)',
                  fontSize: '13px',
                  fontWeight: 600,
                  border: 'none',
                  boxShadow: '0 1px 0 rgba(255,255,255,0.18) inset',
                  cursor: 'pointer',
                }}>
                <Download size={14} />
                Compile
              </button>
            </div>

            {outputFormat === 'mp3' && (
              <div className="flex items-center gap-2">
                <label style={{ color: 'var(--text-muted)', fontSize: '12px' }}>Bitrate</label>
                <select value={mp3Bitrate} onChange={e => setMp3Bitrate(parseInt(e.target.value, 10))}
                  disabled={isCompiling || livePlaying}
                  style={{ padding: '6px 8px', fontSize: '12px' }}>
                  <option value="128">128 kbps · smaller</option>
                  <option value="192">192 kbps · balanced</option>
                  <option value="256">256 kbps · high</option>
                  <option value="320">320 kbps · maximum</option>
                </select>
              </div>
            )}

            {finalAudioUrl && (
              <button onClick={downloadFinal} className="btn btn-secondary">
                <FileAudio size={14} />
                Download {finalAudioFormat?.toUpperCase()} · {formatTime(finalAudioDuration)}
                {finalAudioSize > 0 && <span style={{ color: 'var(--text-dim)', fontWeight: 400 }}>· {formatFileSize(finalAudioSize)}</span>}
              </button>
            )}
          </div>

          {/* Audio quality options */}
          <details className="mt-4 paper hairline" style={{ borderRadius: '8px', padding: '12px 14px' }}>
            <summary className="mono-font text-xs uppercase tracking-wider flex items-center justify-between" style={{ color: 'var(--text-muted)' }}>
              <span>Audio quality settings</span>
              <ChevronDown size={12} />
            </summary>
            <div className="mt-3 grid grid-cols-1 md:grid-cols-2 gap-4">
              <label className="flex items-start gap-2 cursor-pointer">
                <input type="checkbox" checked={normaliseLoudness} onChange={e => setNormaliseLoudness(e.target.checked)}
                  style={{ marginTop: '3px' }} />
                <div className="text-xs" style={{ lineHeight: 1.5 }}>
                  <div className="font-semibold">Normalise loudness</div>
                  <div style={{ color: 'var(--text-muted)' }}>Auto-balance quiet and loud extracts to a consistent level. Recommended.</div>
                </div>
              </label>
              <div>
                <div className="mono-font text-xs uppercase tracking-wider mb-1" style={{ color: 'var(--text-muted)' }}>Crossfade</div>
                <div className="flex items-center gap-2">
                  <select value={crossfadeMs} onChange={e => setCrossfadeMs(parseInt(e.target.value, 10))}
                    style={{ padding: '6px 8px', flex: 1 }}>
                    <option value="0">None — hard cuts</option>
                    <option value="50">Subtle (50ms)</option>
                    <option value="150">Smooth (150ms)</option>
                    <option value="300">Gentle (300ms)</option>
                  </select>
                </div>
                <div className="text-xs mt-1" style={{ color: 'var(--text-muted)', lineHeight: 1.4 }}>
                  Fades each audio segment in/out to prevent clicks and harsh transitions.
                </div>
              </div>
            </div>
          </details>

          {livePlaying && (
            <div className="mt-4 paper hairline p-3" style={{ borderRadius: '2px', background: 'var(--surface)' }}>
              <div className="flex items-center gap-3 mb-2">
                <div className="mono-font text-xs uppercase tracking-wider opacity-60" style={{ minWidth: '60px' }}>
                  Now: {livePaused && <span className="accent">PAUSED</span>}
                </div>
                <div className="text-sm font-semibold flex-1 truncate">{liveCurrentLabel}</div>
                <div className="mono-font text-xs opacity-60">
                  {liveItemIndex + 1} / {liveTotalItems}
                </div>
              </div>
              <div className="h-1" style={{ background: 'var(--surface-elev)', borderRadius: '1px', overflow: 'hidden' }}>
                <div style={{
                  width: `${liveTotalItems > 0 ? ((liveItemIndex + 1) / liveTotalItems) * 100 : 0}%`,
                  height: '100%',
                  background: 'var(--accent)',
                  transition: 'width 0.3s ease',
                }} />
              </div>
            </div>
          )}

          <div className="mt-6 paper hairline p-4" style={{ borderRadius: '2px', background: 'var(--accent-tint)' }}>
            <div className="mono-font text-xs uppercase tracking-wider opacity-60 mb-2 flex items-center gap-1.5">
              <AlertCircle size={12} /> Source compatibility with audio export
            </div>
            <ul className="text-sm leading-relaxed space-y-1 ml-4 list-disc">
              <li><strong>Uploaded audio files</strong> — always exported into the WAV.</li>
              <li><strong>ElevenLabs / OpenAI announcements</strong> — fully exported into the WAV when an API key is provided.</li>
              <li><strong>Browser TTS announcements</strong> — cannot be recorded by the browser; the WAV gets a brief marker tone in their place. Use a paid TTS provider to bake voice into the file.</li>
              <li><strong>YouTube clips</strong> — cannot be exported into the WAV (DRM). They <strong>do</strong> play correctly during "Preview full exam (live)". For a fully-exported file, expand any YouTube clip and use the <code className="mono-font">yt-dlp</code> command to extract the clip locally, then upload it as a file.</li>
              <li><strong>Spotify clips</strong> — full-track playback requires Spotify Premium and works only in live preview (DRM). However: if your clip's <em>end time</em> is within the first 30 seconds of a track <em>and</em> Spotify exposes a 30-second preview for that track, the clip <strong>will</strong> be baked into the file automatically.</li>
            </ul>
          </div>

          <div className="mt-4 hide-mobile">
            <details className="paper hairline" style={{ borderRadius: '8px', padding: '10px 14px' }}>
              <summary className="mono-font text-xs uppercase tracking-wider flex items-center justify-between" style={{ color: 'var(--text-muted)' }}>
                <span>Keyboard shortcuts</span>
                <ChevronDown size={12} />
              </summary>
              <div className="mt-3 grid grid-cols-2 gap-x-6 gap-y-1.5 text-xs" style={{ color: 'var(--text-muted)' }}>
                <div className="flex items-center gap-2"><kbd className="mono-font" style={kbdStyle}>Space</kbd> Play / pause preview</div>
                <div className="flex items-center gap-2"><kbd className="mono-font" style={kbdStyle}>Esc</kbd> Stop preview / close panel</div>
                <div className="flex items-center gap-2"><kbd className="mono-font" style={kbdStyle}>N</kbd> Skip to next extract</div>
                <div className="flex items-center gap-2"><kbd className="mono-font" style={kbdStyle}>P</kbd> Skip to previous extract</div>
                <div className="flex items-center gap-2"><kbd className="mono-font" style={kbdStyle}>K</kbd> Skip current item</div>
                <div className="flex items-center gap-2"><kbd className="mono-font" style={kbdStyle}>T</kbd> Toggle theme</div>
              </div>
            </details>
          </div>
        </section>

        <footer className="mt-16 pt-8 text-center" style={{ borderTop: '0.5px solid var(--border)', fontSize: '11px', color: 'var(--text-faint)' }}>
          Aural Composer
        </footer>
        </main>
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

  const hasAudio = !!q.source;

  return (
    <div className={`question-card ${hasAudio ? 'has-audio' : 'missing-audio'}`}
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
        {/* Left rail - number + handle */}
        <div className="flex flex-col items-center gap-2" style={{
          padding: '20px 10px 16px',
          borderRight: '0.5px solid var(--border)',
          minWidth: '52px',
        }}>
          <div className="cursor-grab" title="Drag to reorder" style={{ color: 'var(--text-faint)' }}>
            <GripVertical size={14} />
          </div>
          <div className="display-font" style={{
            fontSize: '22px',
            fontWeight: 600,
            lineHeight: 1,
            color: 'var(--text)',
            letterSpacing: '-0.02em',
          }}>{index + 1}</div>
          {q.marks != null && (
            <div className="flex items-center gap-0.5" style={{ fontSize: '11px', color: 'var(--text-dim)' }}>
              <Award size={9} /> {q.marks}
            </div>
          )}
          <div className="flex flex-col gap-0 mt-1">
            <button onClick={onMoveUp} disabled={disabled || !onMoveUp} title="Move up"
              className="btn btn-ghost"
              style={{ padding: '2px', minHeight: 0 }}>
              <ChevronUp size={11} />
            </button>
            <button onClick={onMoveDown} disabled={disabled || !onMoveDown} title="Move down"
              className="btn btn-ghost"
              style={{ padding: '2px', minHeight: 0 }}>
              <ChevronDown size={11} />
            </button>
          </div>
        </div>

        {/* Right content area */}
        <div className="flex-1" style={{ padding: '20px 20px 16px', minWidth: 0 }}>
          {/* Title row */}
          <div className="flex items-center gap-2 mb-3">
            <input type="text" value={q.label} onChange={e => onUpdate(q.id, 'label', e.target.value)} disabled={disabled}
              placeholder={`Extract ${index + 1}`}
              className="display-font bg-transparent border-none p-0 flex-1"
              style={{
                fontSize: '17px',
                fontWeight: 600,
                letterSpacing: '-0.01em',
                color: 'var(--text)',
                borderBottom: '0.5px solid transparent',
                minWidth: 0,
              }}
              onFocus={e => e.target.style.borderBottomColor = 'var(--border-strong)'}
              onBlur={e => e.target.style.borderBottomColor = 'transparent'} />
            {!hasAudio && (
              <span className="pill pill-warning hide-mobile" title="No audio attached">
                <AlertCircle size={10} /> Audio needed
              </span>
            )}
            <button onClick={() => onPreview(q)} disabled={!q.source || disabled}
              className={isPreviewing ? "btn btn-primary btn-sm" : "btn btn-secondary btn-sm"}
              title="Preview this extract with announcements and plays">
              {isPreviewing ? <Pause size={12} /> : <Play size={12} />}
              <span className="hide-mobile">Preview</span>
            </button>
            <button onClick={onDelete} disabled={disabled || !onDelete}
              className="btn btn-danger-ghost btn-icon btn-sm" title="Delete extract">
              <Trash2 size={13} />
            </button>
          </div>

          {/* Announcement */}
          <div className="mb-4">
            <label className="field-label">Announcement</label>
            <textarea value={q.intro} onChange={e => onUpdate(q.id, 'intro', e.target.value)} disabled={disabled} rows={2}
              className="w-full" style={{ resize: 'vertical', fontSize: '13px', lineHeight: 1.5 }} />
          </div>

          {/* Timing controls */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
            <div>
              <label className="field-label flex items-center gap-1.5"><Repeat size={11} /> Plays</label>
              <input type="number" min="1" max="10" value={q.plays} onChange={e => onUpdate(q.id, 'plays', Math.max(1, parseInt(e.target.value) || 1))} disabled={disabled} className="w-full" />
            </div>
            <div>
              <label className="field-label flex items-center gap-1.5"><Clock size={11} /> Gap between (sec)</label>
              <input type="number" min="0" value={q.gapBetweenPlays} onChange={e => onUpdate(q.id, 'gapBetweenPlays', Math.max(0, parseInt(e.target.value) || 0))} disabled={disabled} className="w-full" />
            </div>
            <div>
              <label className="field-label flex items-center gap-1.5"><ChevronRight size={11} /> Gap after (sec)</label>
              <input type="number" min="0" value={q.gapAfter} onChange={e => onUpdate(q.id, 'gapAfter', Math.max(0, parseInt(e.target.value) || 0))} disabled={disabled} className="w-full" />
            </div>
            <div>
              <label className="field-label flex items-center gap-1.5"><Award size={11} /> Marks</label>
              <input type="number" min="0" value={q.marks ?? ''} placeholder="—"
                onChange={e => onUpdate(q.id, 'marks', e.target.value === '' ? null : Math.max(0, parseInt(e.target.value) || 0))}
                disabled={disabled} className="w-full" />
            </div>
          </div>

          {/* Audio source area */}
          {q.source ? (
            <div className="drop-zone has-source" style={{ padding: '14px 16px' }}>
              <div className="flex items-center justify-between gap-3">
                <div className="flex items-center gap-3 min-w-0 flex-1">
                  <div style={{
                    width: '32px', height: '32px',
                    background: q.source.kind === 'spotify' ? 'rgba(29,185,84,0.12)' : 'var(--accent-tint-strong)',
                    border: '0.5px solid ' + (q.source.kind === 'spotify' ? 'rgba(29,185,84,0.3)' : 'var(--accent-border)'),
                    borderRadius: 'var(--r-md)',
                    display: 'flex', alignItems: 'center', justifyContent: 'center',
                    flexShrink: 0,
                  }}>
                    {q.source.kind === 'file' && <Music size={15} style={{ color: 'var(--accent-soft)' }} />}
                    {q.source.kind === 'youtube' && <Youtube size={15} style={{ color: 'var(--accent-soft)' }} />}
                    {q.source.kind === 'spotify' && <Music size={15} style={{ color: '#1db954' }} />}
                  </div>
                  <div className="min-w-0">
                    <div style={{ fontSize: '13px', fontWeight: 500, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                      {q.source.kind === 'file' && q.source.name}
                      {q.source.kind === 'youtube' && `YouTube · ${q.source.videoId}`}
                      {q.source.kind === 'spotify' && (
                        <span>{q.source.name} <span style={{ color: 'var(--text-muted)' }}>— {q.source.artists}</span></span>
                      )}
                    </div>
                    <div style={{ fontSize: '11px', color: 'var(--text-dim)', marginTop: '2px' }}>
                      {q.source.kind === 'file' && (() => {
                        const trimStart = q.source.trimStart || 0;
                        const trimEnd = q.source.trimEnd != null ? q.source.trimEnd : q.source.buffer.duration;
                        const playDur = trimEnd - trimStart;
                        const isTrimmed = trimStart > 0 || trimEnd < q.source.buffer.duration - 0.01;
                        return (
                          <>
                            {isTrimmed ? (
                              <>{formatTime(trimStart)} → {formatTime(trimEnd)} · clip {formatTime(playDur)}</>
                            ) : (
                              <>{formatTime(q.source.buffer.duration)}</>
                            )} · {q.plays}× = {formatTime(playDur * q.plays + q.gapBetweenPlays * (q.plays - 1))}
                          </>
                        );
                      })()}
                      {q.source.kind === 'youtube' && (
                        <>{q.source.startStr || formatTime(q.source.start)} → {q.source.endStr || formatTime(q.source.end)} · clip {formatTime(q.source.duration)} · {q.plays}× plays</>
                      )}
                      {q.source.kind === 'spotify' && (
                        <>
                          {q.source.startStr} → {q.source.endStr} · clip {formatTime(q.source.duration)} · {q.plays}× plays
                          {q.source.previewUrl && q.source.end <= 30 && <span className="ml-2" style={{ color: 'var(--success)' }}>· exportable</span>}
                        </>
                      )}
                    </div>
                  </div>
                </div>
                <div className="flex items-center gap-1 flex-shrink-0">
                  {q.source.kind === 'file' && q.source.peaks && (
                    <div className="hide-mobile mr-1">
                      <WaveformThumb
                        peaks={q.source.peaks}
                        trimStart={q.source.trimStart || 0}
                        trimEnd={q.source.trimEnd != null ? q.source.trimEnd : q.source.buffer.duration}
                        totalDuration={q.source.buffer.duration} />
                    </div>
                  )}
                  {q.source.kind === 'spotify' && q.source.externalUrl && (
                    <a href={q.source.externalUrl} target="_blank" rel="noopener" className="btn btn-ghost btn-icon btn-sm" title="Open in Spotify">
                      <ExternalLink size={13} />
                    </a>
                  )}
                  <button onClick={() => onClear(q.id)} disabled={disabled}
                    className="btn btn-danger-ghost btn-icon btn-sm" title="Remove source">
                    <Trash2 size={13} />
                  </button>
                </div>
              </div>
              {q.source.kind === 'youtube' && (
                <details className="mt-3 pt-3" style={{ borderTop: '0.5px solid var(--border)' }}>
                  <summary style={{ fontSize: '12px', color: 'var(--text-muted)', cursor: 'pointer' }}>
                    <span style={{ color: 'var(--text-dim)' }}>▸</span> Convert to local file with yt-dlp
                  </summary>
                  <div className="mt-2 flex items-center gap-2">
                    <code className="mono-font flex-1 overflow-x-auto" style={{ fontSize: '11px', padding: '8px 10px', background: 'var(--code-bg)', color: 'var(--code-text)', borderRadius: 'var(--r-sm)' }}>{ytdlpCommand}</code>
                    <button onClick={() => navigator.clipboard.writeText(ytdlpCommand)}
                      className="btn btn-secondary btn-icon btn-sm" title="Copy">
                      <Copy size={12} />
                    </button>
                  </div>
                  <div style={{ fontSize: '12px', color: 'var(--text-dim)', marginTop: '8px', lineHeight: 1.5 }}>
                    Run this in your terminal to extract just the clip as MP3, then upload it for full audio export support.
                  </div>
                </details>
              )}
              {q.source.kind === 'file' && q.source.buffer && (
                <details className="mt-3 pt-3" style={{ borderTop: '0.5px solid var(--border)' }} open={(q.source.trimStart || 0) > 0 || (q.source.trimEnd != null && q.source.trimEnd < q.source.buffer.duration - 0.01)}>
                  <summary style={{ fontSize: '12px', color: 'var(--text-muted)', cursor: 'pointer' }}>
                    <span style={{ color: 'var(--text-dim)' }}>▸</span> Trim audio
                  </summary>
                  <WaveformTrimmer source={q.source} disabled={disabled}
                    onUpdate={(s, e) => onUpdate(q.id, 'source', { ...q.source, trimStart: s, trimEnd: e })} />
                </details>
              )}
            </div>
          ) : (
            <div>
              <div className="flex gap-1.5 mb-3 flex-wrap">
                <button onClick={() => setMode('file')} className={`tab ${mode === 'file' ? 'active' : ''}`} disabled={disabled}>
                  <Upload size={11} className="inline mr-1" /> File
                </button>
                <button onClick={() => setMode('youtube')} className={`tab ${mode === 'youtube' ? 'active' : ''}`} disabled={disabled}>
                  <Youtube size={11} className="inline mr-1" /> YouTube
                </button>
                <button onClick={() => setMode('spotify')} className={`tab ${mode === 'spotify' ? 'active' : ''}`} disabled={disabled}>
                  <Music size={11} className="inline mr-1" /> Spotify
                </button>
              </div>

              {mode === 'file' && (
                <div className="drop-zone text-center"
                  style={{
                    padding: '20px 16px',
                    background: dragOver ? 'var(--accent-tint-strong)' : 'var(--surface-2)',
                    borderColor: dragOver ? 'var(--accent-border)' : 'var(--border-strong)',
                  }}
                  onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                  onDragLeave={() => setDragOver(false)}
                  onDrop={handleDrop}>
                  <button onClick={() => fileInputRef.current?.click()} disabled={disabled}
                    className="btn btn-ghost"
                    style={{ background: 'transparent', padding: '4px 0' }}>
                    <Upload size={14} /> <span>Drop audio here or click to choose</span>
                  </button>
                  <div style={{ fontSize: '11px', color: 'var(--text-dim)', marginTop: '6px' }}>MP3, WAV, or M4A</div>
                  <input ref={fileInputRef} type="file" accept="audio/*" className="hidden" onChange={e => onFileUpload(q.id, e.target.files[0])} />
                </div>
              )}

              {mode === 'youtube' && (
                <div style={{ background: 'var(--surface-2)', border: '0.5px solid var(--border)', borderRadius: 'var(--r-md)', padding: '14px' }}>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                    <div className="md:col-span-3">
                      <label className="field-label">YouTube URL</label>
                      <input type="text" value={ytUrl} onChange={e => setYtUrl(e.target.value)} placeholder="https://www.youtube.com/watch?v=…" className="w-full" disabled={disabled} />
                    </div>
                    <div>
                      <label className="field-label">Start</label>
                      <input type="text" value={ytStart} onChange={e => setYtStart(e.target.value)} placeholder="0:45 or 1m23s" className="w-full" disabled={disabled} />
                    </div>
                    <div>
                      <label className="field-label">End</label>
                      <input type="text" value={ytEnd} onChange={e => setYtEnd(e.target.value)} placeholder="2:15" className="w-full" disabled={disabled} />
                    </div>
                    <div className="flex items-end">
                      <button onClick={() => {
                        onYouTubeSet(q.id, { url: ytUrl, startStr: ytStart, endStr: ytEnd });
                        setYtUrl(''); setYtStart(''); setYtEnd('');
                      }} disabled={disabled || !ytUrl}
                        className="btn btn-primary btn-sm w-full">
                        Add clip
                      </button>
                    </div>
                  </div>
                  <div style={{ fontSize: '11px', color: 'var(--text-dim)', marginTop: '10px', lineHeight: 1.5 }}>
                    Time formats accepted: <code className="mono-font" style={{ background: 'var(--surface-elev)', padding: '1px 4px', borderRadius: '3px' }}>1:23</code>, <code className="mono-font" style={{ background: 'var(--surface-elev)', padding: '1px 4px', borderRadius: '3px' }}>83</code>, or <code className="mono-font" style={{ background: 'var(--surface-elev)', padding: '1px 4px', borderRadius: '3px' }}>1m23s</code>.
                  </div>
                </div>
              )}

              {mode === 'spotify' && (
                <div style={{ background: 'var(--surface-2)', border: '0.5px solid var(--border)', borderRadius: 'var(--r-md)', padding: '14px' }}>
                  {!spotifyConnected ? (
                    <div className="text-center py-3" style={{ fontSize: '13px', color: 'var(--text-muted)', lineHeight: 1.5 }}>
                      Connect a Spotify account first.<br />
                      <span style={{ color: 'var(--text-dim)' }}>Open <strong style={{ color: 'var(--text-muted)' }}>Voice &amp; API</strong> in the header to connect.</span>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                      <div className="md:col-span-3">
                        <label className="field-label">Spotify track URL or URI</label>
                        <input type="text" value={spUrl} onChange={e => setSpUrl(e.target.value)} placeholder="https://open.spotify.com/track/…" className="w-full" disabled={disabled} />
                      </div>
                      <div>
                        <label className="field-label">Start</label>
                        <input type="text" value={spStart} onChange={e => setSpStart(e.target.value)} placeholder="0:00" className="w-full" disabled={disabled} />
                      </div>
                      <div>
                        <label className="field-label">End</label>
                        <input type="text" value={spEnd} onChange={e => setSpEnd(e.target.value)} placeholder="2:15 (blank = full track)" className="w-full" disabled={disabled} />
                      </div>
                      <div className="flex items-end">
                        <button onClick={async () => {
                          await onSpotifyTrackAdd(q.id, spUrl, spStart, spEnd);
                          setSpUrl(''); setSpStart(''); setSpEnd('');
                        }} disabled={disabled || !spUrl}
                          className="btn btn-primary btn-sm w-full">
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

// ===== MP3 encoder (loads lamejs on demand) =====
let lamejsPromise = null;
function loadLamejs() {
  if (lamejsPromise) return lamejsPromise;
  lamejsPromise = new Promise((resolve, reject) => {
    if (window.lamejs) return resolve(window.lamejs);
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/lamejs@1.2.1/lame.min.js';
    script.onload = () => {
      if (window.lamejs) resolve(window.lamejs);
      else reject(new Error('lamejs loaded but global not found'));
    };
    script.onerror = () => reject(new Error('Failed to load lamejs from CDN'));
    document.head.appendChild(script);
  });
  return lamejsPromise;
}

// ===== lz-string for URL-safe compression of shared configs =====
let lzStringPromise = null;
function loadLzString() {
  if (lzStringPromise) return lzStringPromise;
  lzStringPromise = new Promise((resolve, reject) => {
    if (window.LZString) return resolve(window.LZString);
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/lz-string@1.5.0/libs/lz-string.min.js';
    script.onload = () => {
      if (window.LZString) resolve(window.LZString);
      else reject(new Error('lz-string loaded but global not found'));
    };
    script.onerror = () => reject(new Error('Failed to load lz-string from CDN'));
    document.head.appendChild(script);
  });
  return lzStringPromise;
}

async function audioBufferToMp3(buffer, bitrate = 192, onProgress) {
  const lamejs = await loadLamejs();
  const numChannels = Math.min(buffer.numberOfChannels, 2); // MP3 supports max 2 channels
  const sampleRate = buffer.sampleRate;
  const mp3encoder = new lamejs.Mp3Encoder(numChannels, sampleRate, bitrate);

  // Convert Float32 [-1, 1] samples to Int16 [-32768, 32767]
  const toInt16 = (floatArr) => {
    const out = new Int16Array(floatArr.length);
    for (let i = 0; i < floatArr.length; i++) {
      const s = Math.max(-1, Math.min(1, floatArr[i]));
      out[i] = s < 0 ? s * 0x8000 : s * 0x7FFF;
    }
    return out;
  };

  const left = toInt16(buffer.getChannelData(0));
  const right = numChannels > 1 ? toInt16(buffer.getChannelData(1)) : null;

  const chunkSize = 1152; // standard MP3 frame size
  const mp3Data = [];
  const totalSamples = left.length;
  for (let i = 0; i < totalSamples; i += chunkSize) {
    const leftChunk = left.subarray(i, i + chunkSize);
    const rightChunk = right ? right.subarray(i, i + chunkSize) : null;
    const mp3buf = right
      ? mp3encoder.encodeBuffer(leftChunk, rightChunk)
      : mp3encoder.encodeBuffer(leftChunk);
    if (mp3buf.length > 0) mp3Data.push(mp3buf);
    if (onProgress && i % (chunkSize * 100) === 0) {
      onProgress(i / totalSamples);
      // Yield to the browser to keep the UI responsive
      await new Promise(r => setTimeout(r, 0));
    }
  }
  const flush = mp3encoder.flush();
  if (flush.length > 0) mp3Data.push(flush);

  return new Blob(mp3Data, { type: 'audio/mpeg' });
}

// ===== OGG/WebM encoder using MediaRecorder =====
async function audioBufferToCompressed(buffer, mimeType, bitrate = 128000) {
  return new Promise((resolve, reject) => {
    const ctx = new (window.AudioContext || window.webkitAudioContext)({ sampleRate: buffer.sampleRate });
    const dest = ctx.createMediaStreamDestination();
    const src = ctx.createBufferSource();
    src.buffer = buffer;
    src.connect(dest);

    let recorder;
    try {
      recorder = new MediaRecorder(dest.stream, { mimeType, audioBitsPerSecond: bitrate });
    } catch (e) {
      reject(new Error(`MediaRecorder doesn't support ${mimeType} in this browser`));
      return;
    }

    const chunks = [];
    recorder.ondataavailable = (e) => { if (e.data.size > 0) chunks.push(e.data); };
    recorder.onstop = () => {
      ctx.close();
      resolve(new Blob(chunks, { type: mimeType }));
    };
    recorder.onerror = (e) => { ctx.close(); reject(e.error || new Error('MediaRecorder error')); };

    src.onended = () => {
      // Give the recorder a moment to flush
      setTimeout(() => recorder.stop(), 100);
    };
    recorder.start();
    src.start();
  });
}

function pickSupportedCompressedFormat() {
  const candidates = [
    { mime: 'audio/webm;codecs=opus', ext: 'webm' },
    { mime: 'audio/ogg;codecs=opus', ext: 'ogg' },
    { mime: 'audio/webm', ext: 'webm' },
  ];
  for (const c of candidates) {
    if (MediaRecorder.isTypeSupported(c.mime)) return c;
  }
  return null;
}

// ===== Small inline waveform thumbnail (shown on filled extract cards) =====
// ===== Auth screen (sign in / sign up / reset password) =====
function AuthScreen({ theme, toggleTheme }) {
  const [mode, setMode] = useState('signin'); // 'signin' | 'signup' | 'reset'
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [displayName, setDisplayName] = useState('');
  const [loading, setLoading] = useState(false);
  const [message, setMessage] = useState(null);
  const [error, setError] = useState(null);

  const submit = async (e) => {
    e.preventDefault();
    setError(null);
    setMessage(null);
    setLoading(true);
    try {
      if (mode === 'signin') {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      } else if (mode === 'signup') {
        const { data, error } = await supabase.auth.signUp({
          email,
          password,
          options: { data: { display_name: displayName } },
        });
        if (error) throw error;
        if (!data.session) {
          setMessage('Check your email — we sent a confirmation link. Click it and then come back to sign in.');
        }
      } else if (mode === 'reset') {
        const { error } = await supabase.auth.resetPasswordForEmail(email, {
          redirectTo: window.location.origin + window.location.pathname,
        });
        if (error) throw error;
        setMessage('Password reset email sent. Check your inbox.');
      }
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center px-4 bg-image-auth" style={{
      color: 'var(--text)',
      fontFamily: "'Geist', system-ui, -apple-system, sans-serif",
      paddingTop: '5vh',
      paddingBottom: '5vh',
      position: 'relative',
    }}>

      <button onClick={toggleTheme}
        style={{ position: 'absolute', top: '16px', right: '16px', padding: '8px',
          background: 'var(--surface)', color: 'var(--text-muted)', borderRadius: '8px',
          border: '0.5px solid var(--border)', cursor: 'pointer', zIndex: 2 }}
        title="Toggle theme">
        {theme === 'dark' ? <Sun size={14} /> : <Moon size={14} />}
      </button>

      <div style={{
        background: 'var(--surface-solid)',
        border: '0.5px solid var(--border)',
        borderRadius: '16px',
        padding: '36px 32px',
        width: '100%',
        maxWidth: '400px',
        boxShadow: 'var(--shadow-card)',
        position: 'relative',
        zIndex: 1,
      }}>
        {/* Logo + title */}
        <div className="flex flex-col items-center text-center mb-7">
          <div style={{
            width: '52px', height: '52px',
            background: 'linear-gradient(135deg, var(--accent) 0%, #4f46e5 100%)',
            borderRadius: '13px',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            boxShadow: 'var(--logo-shadow)',
            marginBottom: '16px',
          }}>
            <Music size={24} color="#ffffff" strokeWidth={2.2} />
          </div>
          <h1 className="display-font" style={{ fontSize: '22px', fontWeight: 600, letterSpacing: '-0.02em', margin: 0, marginBottom: '4px' }}>Aural Composer</h1>
          <div className="mono-font" style={{ fontSize: '11px', textTransform: 'uppercase', letterSpacing: '0.1em', color: 'var(--text-dim)' }}>
            {mode === 'signin' && 'Sign in to your account'}
            {mode === 'signup' && 'Create an account'}
            {mode === 'reset' && 'Reset your password'}
          </div>
        </div>

        <form onSubmit={submit} className="space-y-4">
          {mode === 'signup' && (
            <div>
              <label className="mono-font" style={{ fontSize: '10px', textTransform: 'uppercase', letterSpacing: '0.1em', color: 'var(--text-muted)', marginBottom: '6px', display: 'block', fontWeight: 500 }}>Name (optional)</label>
              <input type="text" value={displayName} onChange={e => setDisplayName(e.target.value)}
                placeholder="Rich Holdsworth"
                className="w-full"
                style={{ fontSize: '14px', background: 'var(--input-bg)', border: '0.5px solid var(--border-strong)' }} />
            </div>
          )}
          <div>
            <label className="mono-font" style={{ fontSize: '10px', textTransform: 'uppercase', letterSpacing: '0.1em', color: 'var(--text-muted)', marginBottom: '6px', display: 'block', fontWeight: 500 }}>Email</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)}
              required autoComplete="email"
              placeholder="you@example.com"
              className="w-full"
              style={{ fontSize: '14px', background: 'var(--input-bg)', border: '0.5px solid var(--border-strong)' }} />
          </div>
          {mode !== 'reset' && (
            <div>
              <label className="mono-font" style={{ fontSize: '10px', textTransform: 'uppercase', letterSpacing: '0.1em', color: 'var(--text-muted)', marginBottom: '6px', display: 'block', fontWeight: 500 }}>Password</label>
              <input type="password" value={password} onChange={e => setPassword(e.target.value)}
                required autoComplete={mode === 'signin' ? 'current-password' : 'new-password'}
                minLength={8}
                placeholder="••••••••"
                className="w-full"
                style={{ fontSize: '14px', background: 'var(--input-bg)', border: '0.5px solid var(--border-strong)' }} />
              {mode === 'signup' && (
                <div style={{ fontSize: '11px', marginTop: '6px', color: 'var(--text-dim)' }}>At least 8 characters.</div>
              )}
            </div>
          )}

          {error && (
            <div style={{ fontSize: '12px', padding: '10px 12px', background: 'rgba(220,38,38,0.08)', color: theme === 'dark' ? '#fca5a5' : '#b91c1c', borderRadius: '8px', border: '0.5px solid rgba(220,38,38,0.3)', lineHeight: 1.4 }}>
              {error}
            </div>
          )}
          {message && (
            <div style={{ fontSize: '12px', padding: '10px 12px', background: 'var(--accent-tint)', color: 'var(--accent)', borderRadius: '8px', border: '0.5px solid var(--accent-border)', lineHeight: 1.4 }}>
              {message}
            </div>
          )}

          <button type="submit" disabled={loading}
            className="w-full flex items-center justify-center gap-2 accent-bg"
            style={{
              padding: '11px 16px',
              fontSize: '13px',
              fontWeight: 600,
              borderRadius: '8px',
              border: 'none',
              cursor: loading ? 'not-allowed' : 'pointer',
              marginTop: '4px',
            }}>
            {loading && <Loader2 size={14} className="animate-spin" />}
            {mode === 'signin' && 'Sign in'}
            {mode === 'signup' && 'Create account'}
            {mode === 'reset' && 'Send reset email'}
          </button>
        </form>

        <div style={{ marginTop: '24px', paddingTop: '20px', borderTop: '0.5px solid var(--border)', fontSize: '13px', color: 'var(--text-muted)', lineHeight: 1.7 }}>
          {mode === 'signin' && (
            <>
              <div>Don't have an account? <button type="button" onClick={() => { setMode('signup'); setError(null); setMessage(null); }} style={{ background: 'transparent', border: 'none', color: 'var(--accent)', textDecoration: 'underline', cursor: 'pointer', padding: 0, font: 'inherit' }}>Create one</button></div>
              <div>Forgot your password? <button type="button" onClick={() => { setMode('reset'); setError(null); setMessage(null); }} style={{ background: 'transparent', border: 'none', color: 'var(--accent)', textDecoration: 'underline', cursor: 'pointer', padding: 0, font: 'inherit' }}>Reset it</button></div>
            </>
          )}
          {mode === 'signup' && (
            <div>Already have an account? <button type="button" onClick={() => { setMode('signin'); setError(null); setMessage(null); }} style={{ background: 'transparent', border: 'none', color: 'var(--accent)', textDecoration: 'underline', cursor: 'pointer', padding: 0, font: 'inherit' }}>Sign in</button></div>
          )}
          {mode === 'reset' && (
            <div>Remembered? <button type="button" onClick={() => { setMode('signin'); setError(null); setMessage(null); }} style={{ background: 'transparent', border: 'none', color: 'var(--accent)', textDecoration: 'underline', cursor: 'pointer', padding: 0, font: 'inherit' }}>Sign in</button></div>
          )}
        </div>
      </div>

      <div style={{ marginTop: '20px', fontSize: '12px', color: 'rgba(255,255,255,0.7)', textAlign: 'center', maxWidth: '380px', lineHeight: 1.6, position: 'relative', zIndex: 1, textShadow: '0 1px 4px rgba(0,0,0,0.5)' }}>
        A private workspace for teachers preparing aural and listening exams. New accounts require approval before saving and sharing.
      </div>
    </div>
  );
}

// ===== Pending-approval screen =====
function PendingApprovalScreen({ profile, onSignOut, theme, toggleTheme }) {
  return (
    <div className="min-h-screen flex items-center justify-center px-4 bg-image-auth" style={{
      color: 'var(--text)',
      fontFamily: "'Geist', system-ui, -apple-system, sans-serif",
    }}>
      <button onClick={toggleTheme}
        style={{ position: 'absolute', top: '16px', right: '16px', padding: '8px',
          background: 'var(--surface)', color: 'var(--text-muted)', borderRadius: '8px',
          border: '0.5px solid var(--border)', cursor: 'pointer' }}>
        {theme === 'dark' ? <Sun size={14} /> : <Moon size={14} />}
      </button>

      <div style={{
        background: 'var(--surface-solid)',
        border: '0.5px solid var(--border)',
        borderRadius: '16px',
        padding: '36px 32px',
        width: '100%',
        maxWidth: '440px',
        boxShadow: 'var(--shadow-card)',
        textAlign: 'center',
      }}>
        <div style={{
          width: '56px', height: '56px',
          background: 'var(--accent-tint-strong)',
          borderRadius: '14px',
          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
          marginBottom: '18px',
          border: '0.5px solid var(--accent-border)',
        }}>
          <Mail size={26} className="accent" />
        </div>
        <h1 className="display-font" style={{ fontSize: '22px', fontWeight: 600, letterSpacing: '-0.02em', margin: 0, marginBottom: '10px' }}>Awaiting approval</h1>
        <p style={{ fontSize: '14px', color: 'var(--text-muted)', lineHeight: 1.5, marginBottom: '18px' }}>
          Thanks for signing up, <strong style={{ color: 'var(--text)' }}>{profile.email}</strong>. An admin will approve your account shortly. You'll be able to use the app as soon as they do.
        </p>
        <p style={{ fontSize: '12px', color: 'var(--text-dim)', lineHeight: 1.5, marginBottom: '24px' }}>
          Refresh this page once you've been approved.
        </p>
        <div className="flex flex-col gap-2">
          <button onClick={() => window.location.reload()}
            className="w-full accent-bg flex items-center justify-center gap-2"
            style={{ padding: '11px 16px', fontSize: '13px', fontWeight: 600, borderRadius: '8px', border: 'none', cursor: 'pointer' }}>
            <RefreshCw size={13} />
            Check again
          </button>
          <button onClick={onSignOut}
            className="w-full"
            style={{ padding: '11px 16px', fontSize: '13px', background: 'transparent', color: 'var(--text-muted)', borderRadius: '8px', border: '0.5px solid var(--border)', cursor: 'pointer' }}>
            Sign out
          </button>
        </div>
      </div>
    </div>
  );
}

// ===== Admin: manage user approvals =====
function AdminPanel({ onClose, onChange }) {
  const [profiles, setProfiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  const load = async () => {
    setLoading(true);
    setError(null);
    try {
      const data = await listAllProfiles();
      setProfiles(data);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };
  useEffect(() => { load(); }, []);

  const toggleApproval = async (id, approved) => {
    try {
      await setUserApproved(id, !approved);
      await load();
      onChange?.();
    } catch (err) {
      alert(`Could not update: ${err.message}`);
    }
  };

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 60, display: 'flex', alignItems: 'flex-start', justifyContent: 'center', padding: '40px 16px', overflowY: 'auto' }}
      onClick={onClose}>
      <div className="hairline" style={{ background: 'var(--surface-solid)', borderRadius: '12px', maxWidth: '600px', width: '100%', padding: '20px 24px', boxShadow: 'var(--shadow-card)' }}
        onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-2">
            <Users size={16} className="accent" />
            <h2 className="text-base font-semibold">Manage users</h2>
          </div>
          <button onClick={onClose} className="p-1.5 hairline" style={{ background: 'var(--surface-2)', color: 'var(--text-muted)', borderRadius: '6px' }}>
            <X size={14} />
          </button>
        </div>

        {loading && <div className="text-xs" style={{ color: 'var(--text-muted)' }}><Loader2 size={12} className="animate-spin inline" /> Loading...</div>}
        {error && <div className="text-xs" style={{ color: '#fca5a5' }}>{error}</div>}

        {!loading && profiles.length === 0 && (
          <div className="text-xs" style={{ color: 'var(--text-dim)' }}>No users yet.</div>
        )}

        <div className="space-y-2">
          {profiles.map(p => (
            <div key={p.id} className="hairline p-3 flex items-center justify-between gap-3" style={{ borderRadius: '8px', background: 'var(--surface-2)' }}>
              <div className="min-w-0 flex-1">
                <div className="text-sm font-medium truncate">{p.display_name || p.email}</div>
                <div className="mono-font text-xs truncate" style={{ color: 'var(--text-muted)' }}>{p.email}</div>
                <div className="mono-font text-xs" style={{ color: 'var(--text-dim)' }}>
                  {new Date(p.created_at).toLocaleDateString()} · {p.is_admin && 'Admin · '}{p.approved ? 'Approved' : 'Pending'}
                </div>
              </div>
              <button onClick={() => toggleApproval(p.id, p.approved)}
                disabled={p.is_admin}
                className="px-3 py-1.5 mono-font text-xs uppercase tracking-wider"
                style={{
                  background: p.approved ? 'var(--surface-2)' : 'var(--accent)',
                  color: p.approved ? 'var(--text-muted)' : '#fff',
                  border: '0.5px solid ' + (p.approved ? 'var(--border)' : 'var(--accent)'),
                  borderRadius: '6px',
                  opacity: p.is_admin ? 0.4 : 1,
                  cursor: p.is_admin ? 'not-allowed' : 'pointer',
                }}>
                {p.is_admin ? <ShieldCheck size={11} /> : (p.approved ? 'Revoke' : 'Approve')}
              </button>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ===== Welcome banner (shown to new users until dismissed) =====
function WelcomeBanner({ onDismiss, onOpenHelp }) {
  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 80,
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '16px',
      animation: 'fadeIn 0.2s ease',
    }} onClick={onDismiss}>
      <div onClick={e => e.stopPropagation()} style={{
        background: 'var(--surface-solid)',
        border: '0.5px solid var(--border)',
        borderRadius: '16px',
        padding: '32px 28px',
        maxWidth: '480px',
        width: '100%',
        boxShadow: 'var(--shadow-card)',
      }}>
        <div style={{
          width: '52px', height: '52px',
          background: 'linear-gradient(135deg, var(--accent) 0%, #4f46e5 100%)',
          borderRadius: '13px',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: 'var(--logo-shadow)',
          marginBottom: '16px',
        }}>
          <Sparkles size={24} color="#ffffff" strokeWidth={2.2} />
        </div>
        <h2 className="display-font" style={{ fontSize: '22px', fontWeight: 600, letterSpacing: '-0.02em', margin: 0, marginBottom: '10px' }}>
          Welcome to Aural Composer
        </h2>
        <p style={{ fontSize: '14px', color: 'var(--text-muted)', lineHeight: 1.6, marginBottom: '20px' }}>
          A tool for building listening-exam audio files: stitch together your audio extracts with examiner-style spoken announcements, configure plays, pauses, and answer time, then export as MP3 ready to play in your classroom.
        </p>
        <div style={{ fontSize: '13px', color: 'var(--text-muted)', lineHeight: 1.7, marginBottom: '24px' }}>
          <div className="mono-font" style={{ fontSize: '10px', textTransform: 'uppercase', letterSpacing: '0.1em', color: 'var(--text-dim)', marginBottom: '8px' }}>How it works</div>
          <ol style={{ paddingLeft: '20px', margin: 0 }}>
            <li>Drop in your exam paper PDF (or use the default Trinity setup)</li>
            <li>Add audio to each extract — upload, YouTube link, or Spotify track</li>
            <li>Preview the whole exam to check timing and announcements</li>
            <li>Compile and download as MP3, WAV, or OGG</li>
          </ol>
        </div>
        <div style={{ display: 'flex', gap: '8px' }}>
          <button onClick={onDismiss}
            className="accent-bg flex-1"
            style={{ padding: '11px 16px', fontSize: '13px', fontWeight: 600, borderRadius: '8px', border: 'none', cursor: 'pointer' }}>
            Get started
          </button>
          <button onClick={onOpenHelp}
            style={{ padding: '11px 16px', fontSize: '13px', background: 'transparent', color: 'var(--text-muted)', borderRadius: '8px', border: '0.5px solid var(--border)', cursor: 'pointer' }}>
            See full guide
          </button>
        </div>
      </div>
    </div>
  );
}

// ===== Help panel (?  button in header) =====
function HelpPanel({ onClose }) {
  const [section, setSection] = useState('overview');
  const sections = [
    { id: 'overview', label: 'Overview' },
    { id: 'workflow', label: 'Workflow' },
    { id: 'audio', label: 'Audio sources' },
    { id: 'announcements', label: 'Announcements' },
    { id: 'sharing', label: 'Saving & sharing' },
    { id: 'shortcuts', label: 'Shortcuts' },
    { id: 'troubleshooting', label: 'Troubleshooting' },
  ];

  return (
    <div style={{
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 70,
      display: 'flex', alignItems: 'flex-start', justifyContent: 'center',
      padding: '40px 16px', overflowY: 'auto',
    }} onClick={onClose}>
      <div onClick={e => e.stopPropagation()} style={{
        background: 'var(--surface-solid)',
        border: '0.5px solid var(--border)',
        borderRadius: '14px',
        maxWidth: '720px',
        width: '100%',
        boxShadow: 'var(--shadow-card)',
        display: 'flex',
        flexDirection: 'column',
        maxHeight: 'calc(100vh - 80px)',
      }}>
        {/* Header */}
        <div style={{ padding: '20px 24px', borderBottom: '0.5px solid var(--border)', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div className="flex items-center gap-2">
            <HelpCircle size={16} className="accent" />
            <h2 className="display-font" style={{ fontSize: '17px', fontWeight: 600, margin: 0 }}>How to use Aural Composer</h2>
          </div>
          <button onClick={onClose} className="p-1.5 hairline" style={{ background: 'var(--surface-2)', color: 'var(--text-muted)', borderRadius: '6px' }}>
            <X size={14} />
          </button>
        </div>

        {/* Section tabs */}
        <div style={{ padding: '12px 24px', borderBottom: '0.5px solid var(--border)', overflowX: 'auto' }}>
          <div className="flex gap-1.5" style={{ minWidth: 'max-content' }}>
            {sections.map(s => (
              <button key={s.id} onClick={() => setSection(s.id)}
                className="mono-font text-xs uppercase tracking-wider"
                style={{
                  padding: '6px 12px',
                  background: section === s.id ? 'var(--accent)' : 'transparent',
                  color: section === s.id ? 'var(--accent-bg-on)' : 'var(--text-muted)',
                  border: '0.5px solid ' + (section === s.id ? 'var(--accent)' : 'var(--border)'),
                  borderRadius: '6px',
                  cursor: 'pointer',
                  whiteSpace: 'nowrap',
                }}>
                {s.label}
              </button>
            ))}
          </div>
        </div>

        {/* Section content */}
        <div style={{ padding: '24px', overflowY: 'auto', fontSize: '14px', lineHeight: 1.6, color: 'var(--text)' }}>
          {section === 'overview' && (
            <div>
              <p style={{ marginTop: 0 }}>
                Aural Composer builds the listening section of a music or language exam. You provide the audio extracts, and the app assembles them into a single timed file with examiner-style spoken announcements between each extract — "Question 1. You will hear this extract three times." — plus the gaps, repeats, and answer time you'd hear in a real exam.
              </p>
              <p>
                The compiled file plays end-to-end on any device with no human intervention: ideal for classroom playback, exam centres without a CD player, or sending to candidates remotely.
              </p>
              <p style={{ marginBottom: 0, color: 'var(--text-muted)', fontSize: '13px' }}>
                Everything you set up — extracts, announcements, timings — is saved automatically to your account. Sign in from any device to pick up where you left off.
              </p>
            </div>
          )}

          {section === 'workflow' && (
            <div>
              <ol style={{ paddingLeft: '20px', margin: 0 }}>
                <li style={{ marginBottom: '14px' }}>
                  <strong>Start from an exam paper PDF</strong> — drop it on the "Start from a paper" zone. The app auto-detects extract headings, play counts, and marks, and creates the right number of extracts for you. Or just edit the default Trinity setup below.
                </li>
                <li style={{ marginBottom: '14px' }}>
                  <strong>Add audio to each extract</strong> — upload a file, paste a YouTube link with start/end timestamps, or import a Spotify track. You can drag-reorder extracts, edit announcement text, and trim audio clips precisely.
                </li>
                <li style={{ marginBottom: '14px' }}>
                  <strong>Set timings</strong> — for each extract: how many plays, the gap between plays (thinking time), and the gap after (answer time before the next extract starts). Defaults match Trinity GCSE timings.
                </li>
                <li style={{ marginBottom: '14px' }}>
                  <strong>Preview the full exam</strong> — click "Preview full exam (live)" to play through everything in real time. You can pause, skip extracts, or skip individual segments using the transport controls.
                </li>
                <li style={{ marginBottom: 0 }}>
                  <strong>Compile and download</strong> — choose MP3, WAV, or OGG, click Compile, wait a minute or so, and download the file.
                </li>
              </ol>
            </div>
          )}

          {section === 'audio' && (
            <div>
              <p style={{ marginTop: 0 }}>
                Three sources are supported, with different export characteristics:
              </p>
              <ul style={{ paddingLeft: '20px', margin: 0 }}>
                <li style={{ marginBottom: '12px' }}>
                  <strong>Uploaded audio files</strong> (MP3, WAV, M4A) — best option. Always exported into the compiled file. Can be trimmed precisely using the waveform editor that appears when you expand an extract.
                </li>
                <li style={{ marginBottom: '12px' }}>
                  <strong>YouTube clips</strong> — paste any YouTube URL with start and end timestamps. Plays correctly during live preview, but YouTube blocks third-party apps from including its content in downloads, so the clip will be silent in the compiled file. If you need it in the compiled file, expand the extract — the app shows you a <code className="mono-font" style={{ background: 'var(--surface-elev)', padding: '1px 5px', borderRadius: '3px' }}>yt-dlp</code> command. This is a free command-line tool; if you're not familiar with terminals, ask a tech-comfortable colleague. Once you have the clip as an MP3, upload it normally.
                </li>
                <li style={{ marginBottom: 0 }}>
                  <strong>Spotify tracks</strong> — use a Spotify track as your audio source. You'll need to connect a Spotify account in Voice & API. Two important caveats: (1) full-track playback only works in live preview, not in the compiled file, because Spotify forbids redistribution. (2) Full-track playback also requires a Spotify Premium account. The one exception: if your clip's end time falls within the first 30 seconds of the track, the app can use Spotify's free 30-second preview and bake those clips into the compiled file.
                </li>
              </ul>
            </div>
          )}

          {section === 'announcements' && (
            <div>
              <p style={{ marginTop: 0 }}>
                Spoken examiner announcements introduce each extract, count down between plays, and bookend the exam. By default the app uses your browser's built-in speech synthesis (free, but can't be recorded into the compiled file — you'll get short marker tones in the export instead).
              </p>
              <p>
                For voice baked into the compiled file, you need an API key from either:
              </p>
              <ul style={{ paddingLeft: '20px' }}>
                <li style={{ marginBottom: '8px' }}><strong>OpenAI</strong> (cheaper, very natural voices including British-accent "Fable"). About $0.015 per minute of speech. Sign up at platform.openai.com — costs about $5 minimum credit.</li>
                <li style={{ marginBottom: 0 }}><strong>ElevenLabs</strong> (more premium, natural British voices). Higher quality but more expensive per minute.</li>
              </ul>
              <p>
                Set this up under <strong>Voice & API</strong> in the header. Keys are stored only in your browser — not on any server.
              </p>
              <p style={{ marginBottom: 0 }}>
                You can edit the announcement text for each extract individually, or change the standard phrasing (opening, between-plays, closing) in the <strong>Announcement script</strong> panel.
              </p>
            </div>
          )}

          {section === 'sharing' && (
            <div>
              <p style={{ marginTop: 0 }}>
                Saved exams have three layers:
              </p>
              <ul style={{ paddingLeft: '20px' }}>
                <li style={{ marginBottom: '12px' }}>
                  <strong>Cloud workspace</strong> (default when signed in) — saved to your account, available from any device. Marked with a cloud icon in the sidebar. You can share specific exams with all approved workspace members using the dropdown menu on each exam.
                </li>
                <li style={{ marginBottom: '12px' }}>
                  <strong>Browser-local</strong> — exams saved before you had an account, or if cloud save fails. Stays in this browser only. Marked LOCAL with a strikethrough-cloud icon.
                </li>
                <li style={{ marginBottom: 0 }}>
                  <strong>JSON config files / share links</strong> — under Save config / Load config in the toolbar. Use this to back up your exam, email it to a colleague who hasn't joined the workspace, or share a link that loads the exam (audio files not included).
                </li>
              </ul>
            </div>
          )}

          {section === 'shortcuts' && (
            <div>
              <p style={{ marginTop: 0 }}>Keyboard shortcuts (when not typing in a text field):</p>
              <div style={{ display: 'grid', gridTemplateColumns: '80px 1fr', gap: '8px 16px', marginTop: '12px' }}>
                <kbd className="mono-font" style={kbdStyle}>Space</kbd><div>Play or pause the live preview</div>
                <kbd className="mono-font" style={kbdStyle}>Esc</kbd><div>Stop preview, or close any open panel</div>
                <kbd className="mono-font" style={kbdStyle}>N</kbd><div>Skip forward to the next extract</div>
                <kbd className="mono-font" style={kbdStyle}>P</kbd><div>Skip back to the previous extract</div>
                <kbd className="mono-font" style={kbdStyle}>K</kbd><div>Skip the current item (announcement, silence, or audio segment)</div>
                <kbd className="mono-font" style={kbdStyle}>T</kbd><div>Toggle between dark and light theme</div>
              </div>
            </div>
          )}

          {section === 'troubleshooting' && (
            <div>
              <div style={{ marginBottom: '16px' }}>
                <strong>The compiled file is silent in places.</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--text-muted)' }}>
                  YouTube and Spotify clips can't be baked into the compiled file (DRM). Use uploaded MP3 files for full export support. For YouTube clips, expand the extract and copy the <code className="mono-font">yt-dlp</code> command to extract the clip locally.
                </p>
              </div>
              <div style={{ marginBottom: '16px' }}>
                <strong>Announcements are silent or replaced by short beeps.</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--text-muted)' }}>
                  Browser TTS can't be recorded into the compiled file. Set up OpenAI or ElevenLabs under Voice & API to get spoken announcements baked into the audio.
                </p>
              </div>
              <div style={{ marginBottom: '16px' }}>
                <strong>The compile takes a long time.</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--text-muted)' }}>
                  MP3 encoding runs in the browser and can take 30 seconds to 2 minutes depending on length. WAV is instant but produces large files. OGG is fast but has less universal support.
                </p>
              </div>
              <div style={{ marginBottom: '16px' }}>
                <strong>Audio levels are inconsistent between extracts.</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--text-muted)' }}>
                  Enable "Normalise loudness" under Audio quality settings before compiling. The app will balance quiet and loud extracts to a consistent level.
                </p>
              </div>
              <div style={{ marginBottom: 0 }}>
                <strong>Something else is wrong.</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--text-muted)' }}>
                  Email <a href="mailto:rholdsworth82@gmail.com" style={{ color: 'var(--accent)' }}>rholdsworth82@gmail.com</a> with a screenshot if possible — describe what you were trying to do and what happened instead.
                </p>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// Compact stat cell for the exam dashboard
function StatCell({ label, value, icon, accent }) {
  return (
    <div className="flex items-center gap-2" style={{ paddingTop: '4px', paddingBottom: '4px' }}>
      <span style={{ fontSize: '12px', color: 'var(--text-dim)' }}>{label}</span>
      <span className="flex items-center gap-1" style={{
        fontSize: '15px',
        fontWeight: 600,
        color: accent ? 'var(--accent-soft)' : 'var(--text)',
        letterSpacing: '-0.01em',
      }}>
        {icon && <span style={{ color: 'var(--text-dim)' }}>{icon}</span>}
        {value}
      </span>
    </div>
  );
}

function WaveformThumb({ peaks, trimStart, trimEnd, totalDuration, width = 120, height = 22 }) {
  if (!peaks || peaks.length === 0) return null;
  const startFrac = totalDuration ? (trimStart || 0) / totalDuration : 0;
  const endFrac = totalDuration ? (trimEnd != null ? trimEnd : totalDuration) / totalDuration : 1;
  const startX = startFrac * width;
  const endX = endFrac * width;
  const barWidth = Math.max(1, width / peaks.length - 0.5);
  const stride = width / peaks.length;
  const mid = height / 2;
  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ flexShrink: 0 }}>
      {peaks.map((peak, i) => {
        const x = i * stride;
        const inside = x >= startX - 0.5 && x <= endX + 0.5;
        const h = Math.max(1, peak * (height - 2));
        return (
          <rect
            key={i}
            x={x}
            y={mid - h / 2}
            width={barWidth}
            height={h}
            fill={inside ? 'var(--accent-soft)' : 'var(--text-faint)'}
            opacity={inside ? 0.85 : 0.5}
          />
        );
      })}
    </svg>
  );
}

// ===== Saved exam row in sidebar =====
function SavedExamRow({ entry, onLoad, onUpdate, onRename, onDelete, onToggleShare }) {
  const [menuOpen, setMenuOpen] = useState(false);
  const date = new Date(entry.savedAt);
  const dateStr = date.toLocaleDateString(undefined, { day: 'numeric', month: 'short' });
  const extractCount = entry.config?.questions?.length || 0;
  const ownerName = entry.ownerName || entry.ownerEmail;

  return (
    <div className="group relative" style={{ borderRadius: 'var(--r-md)' }}>
      <button onClick={onLoad}
        className="w-full text-left hover-glow"
        style={{ background: 'transparent', borderRadius: 'var(--r-md)', padding: '8px 10px', border: 'none', cursor: 'pointer' }}
        title={`Load "${entry.name}"`}>
        <div className="flex items-center gap-2" style={{ paddingRight: '20px' }}>
          {entry.kind === 'cloud'
            ? <Cloud size={11} style={{ flexShrink: 0, color: 'var(--accent-soft)' }} />
            : <CloudOff size={11} style={{ flexShrink: 0, color: 'var(--text-faint)' }} />
          }
          <div style={{ fontSize: '13px', fontWeight: 500, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{entry.name}</div>
        </div>
        <div className="flex items-center gap-1.5 mt-0.5" style={{ paddingLeft: '19px', fontSize: '11px', color: 'var(--text-dim)' }}>
          <span>{extractCount} extract{extractCount === 1 ? '' : 's'}</span>
          <span style={{ color: 'var(--text-faint)' }}>·</span>
          <span>{dateStr}</span>
          {entry.kind === 'cloud' && entry.shared_with_all && (
            <span className="pill pill-accent" style={{ fontSize: '10px', padding: '1px 6px', marginLeft: '2px' }}>Shared</span>
          )}
        </div>
        {entry.kind === 'cloud' && !entry.isMine && ownerName && (
          <div className="mt-0.5" style={{ paddingLeft: '19px', fontSize: '11px', color: 'var(--text-faint)' }}>
            by {ownerName}
          </div>
        )}
      </button>
      {(onUpdate || onRename || onDelete || onToggleShare) && (
        <button onClick={(e) => { e.stopPropagation(); setMenuOpen(!menuOpen); }}
          className="absolute top-1.5 right-1.5 opacity-0 group-hover:opacity-100"
          style={{ background: 'transparent', borderRadius: '4px', padding: '4px', border: 'none', cursor: 'pointer', color: 'var(--text-muted)', transition: 'opacity 0.15s' }}
          title="Options">
          <ChevronDown size={12} />
        </button>
      )}

      {menuOpen && (
        <>
          <div onClick={() => setMenuOpen(false)}
            style={{ position: 'fixed', inset: 0, zIndex: 10 }} />
          <div className="absolute right-1 top-9"
            style={{
              background: 'var(--surface-solid)',
              border: '0.5px solid var(--border)',
              borderRadius: 'var(--r-md)',
              minWidth: '180px',
              zIndex: 11,
              boxShadow: 'var(--shadow-lg)',
              overflow: 'hidden',
            }}>
            {onUpdate && (
              <button onClick={() => { setMenuOpen(false); onUpdate(); }}
                className="w-full text-left hover-glow flex items-center gap-2"
                style={{ background: 'transparent', border: 'none', padding: '8px 12px', fontSize: '12px', color: 'var(--text)', cursor: 'pointer' }}>
                <Save size={11} /> Update with current
              </button>
            )}
            {onToggleShare && (
              <button onClick={() => { setMenuOpen(false); onToggleShare(); }}
                className="w-full text-left hover-glow flex items-center gap-2"
                style={{ background: 'transparent', border: 'none', padding: '8px 12px', fontSize: '12px', color: 'var(--text)', cursor: 'pointer' }}>
                {entry.shared_with_all ? <CloudOff size={11} /> : <Cloud size={11} />}
                {entry.shared_with_all ? 'Stop sharing' : 'Share with workspace'}
              </button>
            )}
            {onRename && (
              <button onClick={() => { setMenuOpen(false); onRename(); }}
                className="w-full text-left hover-glow flex items-center gap-2"
                style={{ background: 'transparent', border: 'none', padding: '8px 12px', fontSize: '12px', color: 'var(--text)', cursor: 'pointer' }}>
                <FileText size={11} /> Rename
              </button>
            )}
            {onDelete && (
              <button onClick={() => { setMenuOpen(false); onDelete(); }}
                className="w-full text-left flex items-center gap-2"
                style={{ background: 'transparent', border: 'none', padding: '8px 12px', fontSize: '12px', color: 'var(--danger)', cursor: 'pointer', borderTop: '0.5px solid var(--border)' }}
                onMouseEnter={(e) => e.currentTarget.style.background = 'var(--danger-tint)'}
                onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}>
                <Trash2 size={11} /> Delete
              </button>
            )}
          </div>
        </>
      )}
    </div>
  );
}

// ===== Waveform trimmer =====
function WaveformTrimmer({ source, onUpdate, disabled }) {
  const canvasRef = useRef(null);
  const containerRef = useRef(null);
  const [containerWidth, setContainerWidth] = useState(600);
  const buffer = source.buffer;
  const totalDur = buffer.duration;

  const [trimStart, setTrimStart] = useState(source.trimStart || 0);
  const [trimEnd, setTrimEnd] = useState(source.trimEnd != null ? source.trimEnd : totalDur);
  const [dragging, setDragging] = useState(null); // 'start' | 'end' | null
  const [playhead, setPlayhead] = useState(null); // seconds, while playing
  const [isPlaying, setIsPlaying] = useState(false);
  const playSourceRef = useRef(null);
  const playStartRef = useRef(0);
  const playRafRef = useRef(null);

  // Width tracking
  useEffect(() => {
    if (!containerRef.current) return;
    const ro = new ResizeObserver(entries => {
      for (const e of entries) setContainerWidth(Math.floor(e.contentRect.width));
    });
    ro.observe(containerRef.current);
    return () => ro.disconnect();
  }, []);

  // Draw waveform
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    const W = containerWidth;
    const H = 80;
    canvas.width = W * dpr;
    canvas.height = H * dpr;
    canvas.style.width = W + 'px';
    canvas.style.height = H + 'px';
    const ctx = canvas.getContext('2d');
    ctx.scale(dpr, dpr);
    ctx.clearRect(0, 0, W, H);

    // Compute peaks
    const data = buffer.getChannelData(0);
    const samplesPerPixel = Math.max(1, Math.floor(data.length / W));
    const mid = H / 2;

    // Draw inactive (outside trim) in muted, active (inside trim) in accent
    const startX = (trimStart / totalDur) * W;
    const endX = (trimEnd / totalDur) * W;

    for (let x = 0; x < W; x++) {
      let min = 0, max = 0;
      const offset = x * samplesPerPixel;
      for (let i = 0; i < samplesPerPixel; i++) {
        const v = data[offset + i] || 0;
        if (v < min) min = v;
        if (v > max) max = v;
      }
      const yMin = mid + min * mid * 0.85;
      const yMax = mid + max * mid * 0.85;
      const inside = x >= startX && x <= endX;
      ctx.fillStyle = inside ? 'var(--accent)' : 'rgba(255,255,255,0.15)';
      ctx.fillRect(x, yMin, 1, Math.max(1, yMax - yMin));
    }

    // Center line
    ctx.fillStyle = 'var(--border)';
    ctx.fillRect(0, mid, W, 1);
  }, [containerWidth, buffer, trimStart, trimEnd, totalDur]);

  // Mouse / touch handlers
  const pixelToTime = (px) => Math.max(0, Math.min(totalDur, (px / containerWidth) * totalDur));

  const onPointerDown = (e, which) => {
    if (disabled) return;
    e.preventDefault();
    setDragging(which);
  };

  useEffect(() => {
    if (!dragging) return;
    const onMove = (e) => {
      if (!containerRef.current) return;
      const rect = containerRef.current.getBoundingClientRect();
      const x = (e.touches ? e.touches[0].clientX : e.clientX) - rect.left;
      const t = pixelToTime(x);
      if (dragging === 'start') {
        const newStart = Math.min(t, trimEnd - 0.5);
        setTrimStart(Math.max(0, newStart));
      } else if (dragging === 'end') {
        const newEnd = Math.max(t, trimStart + 0.5);
        setTrimEnd(Math.min(totalDur, newEnd));
      }
    };
    const onUp = () => setDragging(null);
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    window.addEventListener('touchmove', onMove);
    window.addEventListener('touchend', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      window.removeEventListener('touchmove', onMove);
      window.removeEventListener('touchend', onUp);
    };
  }, [dragging, trimEnd, trimStart, totalDur, containerWidth]);

  // Commit values to parent when drag ends
  useEffect(() => {
    if (dragging) return;
    if ((trimStart !== (source.trimStart || 0)) || (trimEnd !== (source.trimEnd != null ? source.trimEnd : totalDur))) {
      onUpdate(trimStart, trimEnd);
    }
  }, [dragging, trimStart, trimEnd]);

  // Reset trim if source changes
  useEffect(() => {
    setTrimStart(source.trimStart || 0);
    setTrimEnd(source.trimEnd != null ? source.trimEnd : totalDur);
  }, [source.name, totalDur]);

  // Playback of trimmed selection
  const stopPlayback = () => {
    if (playSourceRef.current) {
      try { playSourceRef.current.stop(); } catch (e) {}
      playSourceRef.current = null;
    }
    if (playRafRef.current) {
      cancelAnimationFrame(playRafRef.current);
      playRafRef.current = null;
    }
    setIsPlaying(false);
    setPlayhead(null);
  };

  const startPlayback = (fromTime = null) => {
    stopPlayback();
    const startAt = fromTime != null ? fromTime : trimStart;
    const ctx = new (window.AudioContext || window.webkitAudioContext)();
    const src = ctx.createBufferSource();
    src.buffer = buffer;
    src.connect(ctx.destination);
    const playDur = trimEnd - startAt;
    src.start(0, startAt, playDur);
    playSourceRef.current = src;
    playStartRef.current = ctx.currentTime - 0; // time elapsed = ctx.currentTime
    const baseAudioTime = ctx.currentTime;
    setIsPlaying(true);
    setPlayhead(startAt);
    const tick = () => {
      const elapsed = ctx.currentTime - baseAudioTime;
      const current = startAt + elapsed;
      if (current >= trimEnd) {
        stopPlayback();
        return;
      }
      setPlayhead(current);
      playRafRef.current = requestAnimationFrame(tick);
    };
    playRafRef.current = requestAnimationFrame(tick);
    src.onended = () => stopPlayback();
  };

  useEffect(() => () => stopPlayback(), []);

  const playheadX = playhead != null ? (playhead / totalDur) * containerWidth : null;
  const startX = (trimStart / totalDur) * containerWidth;
  const endX = (trimEnd / totalDur) * containerWidth;

  return (
    <div className="mt-3 p-3 hairline" style={{ borderRadius: '2px', background: 'var(--waveform-bg)' }}>
      <div className="flex items-center justify-between mb-2">
        <div className="mono-font text-xs uppercase tracking-wider opacity-60">Trim</div>
        <div className="mono-font text-xs opacity-70">
          {formatTime(trimStart)} → {formatTime(trimEnd)} · clip {formatTime(trimEnd - trimStart)}
        </div>
      </div>

      <div ref={containerRef} style={{ position: 'relative', height: '80px', userSelect: 'none', cursor: 'crosshair' }}>
        <canvas ref={canvasRef} style={{ display: 'block', width: '100%', height: '80px' }} />

        {/* Start handle */}
        <div
          onMouseDown={(e) => onPointerDown(e, 'start')}
          onTouchStart={(e) => onPointerDown(e, 'start')}
          style={{
            position: 'absolute', top: 0, left: `${startX}px`, transform: 'translateX(-50%)',
            width: '12px', height: '80px', cursor: 'ew-resize',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
          <div style={{ width: '3px', height: '100%', background: 'var(--accent)' }} />
          <div style={{
            position: 'absolute', top: '-2px', width: '12px', height: '12px',
            background: 'var(--accent)', borderRadius: '2px',
          }} />
        </div>

        {/* End handle */}
        <div
          onMouseDown={(e) => onPointerDown(e, 'end')}
          onTouchStart={(e) => onPointerDown(e, 'end')}
          style={{
            position: 'absolute', top: 0, left: `${endX}px`, transform: 'translateX(-50%)',
            width: '12px', height: '80px', cursor: 'ew-resize',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
          <div style={{ width: '3px', height: '100%', background: 'var(--accent)' }} />
          <div style={{
            position: 'absolute', bottom: '-2px', width: '12px', height: '12px',
            background: 'var(--accent)', borderRadius: '2px',
          }} />
        </div>

        {/* Playhead */}
        {playheadX != null && (
          <div style={{
            position: 'absolute', top: 0, left: `${playheadX}px`,
            width: '2px', height: '80px', background: 'var(--text)',
            pointerEvents: 'none', boxShadow: '0 0 4px rgba(0,0,0,0.3)',
          }} />
        )}
      </div>

      {/* Controls */}
      <div className="flex items-center gap-2 mt-3">
        <button
          onClick={() => isPlaying ? stopPlayback() : startPlayback()}
          disabled={disabled}
          className="flex items-center gap-1 px-3 py-1.5 hairline mono-font text-xs uppercase tracking-wider"
          style={{ background: isPlaying ? 'var(--accent)' : 'transparent', color: isPlaying ? 'var(--surface)' : 'inherit' }}>
          {isPlaying ? <Pause size={12} /> : <Play size={12} />}
          {isPlaying ? 'Stop' : 'Play selection'}
        </button>

        <div className="flex items-center gap-1.5">
          <label className="mono-font text-xs uppercase tracking-wider opacity-60">Start</label>
          <input type="text" value={formatTime(trimStart)}
            onChange={(e) => {
              const t = parseTimestamp(e.target.value);
              if (t < trimEnd - 0.5) setTrimStart(Math.max(0, t));
            }}
            onBlur={() => onUpdate(trimStart, trimEnd)}
            disabled={disabled}
            className="w-16 text-xs mono-font" style={{ padding: '4px 6px' }} />
        </div>

        <div className="flex items-center gap-1.5">
          <label className="mono-font text-xs uppercase tracking-wider opacity-60">End</label>
          <input type="text" value={formatTime(trimEnd)}
            onChange={(e) => {
              const t = parseTimestamp(e.target.value);
              if (t > trimStart + 0.5) setTrimEnd(Math.min(totalDur, t));
            }}
            onBlur={() => onUpdate(trimStart, trimEnd)}
            disabled={disabled}
            className="w-16 text-xs mono-font" style={{ padding: '4px 6px' }} />
        </div>

        <button
          onClick={() => { setTrimStart(0); setTrimEnd(totalDur); }}
          disabled={disabled || (trimStart === 0 && trimEnd === totalDur)}
          className="mono-font text-xs uppercase tracking-wider opacity-60 hover:opacity-100 underline"
          style={{ background: 'transparent' }}>
          Reset
        </button>

        <div className="flex-1" />
        <div className="mono-font text-xs opacity-50">full: {formatTime(totalDur)}</div>
      </div>
    </div>
  );
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
        background: dragOver ? 'var(--accent-tint)' : 'transparent',
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
    <div className="flex items-center gap-3 py-2 px-3 hairline" style={{ borderRadius: '2px', background: 'var(--surface)' }}>
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
        style={{ color: 'var(--surface)', borderRadius: '2px' }}>
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
