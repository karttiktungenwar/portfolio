// SERVER-ONLY module. Never import this from a "use client" component —
// doing so would attempt to bundle firebase-admin (and your service
// account credentials) into client-side JavaScript.
//
// This file is only ever imported from files under app/api/**/route.ts,
// which run exclusively on the server (Node.js runtime), so the
// environment variables below are never sent to the browser.

import { cert, getApps, initializeApp, type App } from "firebase-admin/app";
import { getFirestore } from "firebase-admin/firestore";

function getFirebaseAdminApp(): App {
  const existing = getApps();
  if (existing.length > 0) {
    return existing[0];
  }

  const projectId = process.env.FIREBASE_PROJECT_ID;
  const clientEmail = process.env.FIREBASE_CLIENT_EMAIL;
  // Private keys in .env files store literal "\n" sequences instead of
  // real newlines — this restores them before handing the key to Firebase.
  const privateKey = process.env.FIREBASE_PRIVATE_KEY?.replace(/\\n/g, "\n");

  if (!projectId || !clientEmail || !privateKey) {
    throw new Error(
      "Missing Firebase Admin credentials. Set FIREBASE_PROJECT_ID, " +
        "FIREBASE_CLIENT_EMAIL and FIREBASE_PRIVATE_KEY in your .env.local " +
        "(or in your hosting provider's environment variable settings)."
    );
  }

  return initializeApp({
    credential: cert({ projectId, clientEmail, privateKey })
  });
}

export function getAdminFirestore() {
  const app = getFirebaseAdminApp();
  return getFirestore(app);
}
