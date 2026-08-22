import { useState, useEffect } from 'react';
import { doc, updateDoc, arrayUnion, arrayRemove } from 'firebase/firestore';
import { db } from '../firebase';

function urlBase64ToUint8Array(base64String: string) {
  const padding = '='.repeat((4 - base64String.length % 4) % 4);
  const base64 = (base64String + padding)
    .replace(/\-/g, '+')
    .replace(/_/g, '/');

  const rawData = window.atob(base64);
  const outputArray = new Uint8Array(rawData.length);

  for (let i = 0; i < rawData.length; ++i) {
    outputArray[i] = rawData.charCodeAt(i);
  }
  return outputArray;
}

export function usePushNotifications(userId: string | undefined) {
  const [isSupported, setIsSupported] = useState(false);
  const [isSubscribed, setIsSubscribed] = useState(false);
  const [permission, setPermission] = useState(Notification.permission);

  useEffect(() => {
    if ('serviceWorker' in navigator && 'PushManager' in window) {
      setIsSupported(true);
      // Register service worker if not already
      navigator.serviceWorker.register('/sw.js').catch(err => {
        console.error('Service Worker registration failed:', err);
      });
      checkSubscription();
    }
  }, []);

  const checkSubscription = async () => {
    try {
      const registration = await navigator.serviceWorker.ready;
      const subscription = await registration.pushManager.getSubscription();
      setIsSubscribed(!!subscription);
    } catch (e) {
      console.error('Error checking subscription', e);
    }
  };

  const subscribeToPush = async () => {
    if (!userId) return;
    try {
      const permissionResult = await Notification.requestPermission();
      setPermission(permissionResult);
      if (permissionResult !== 'granted') {
        throw new Error('Bạn phân quyền từ chối thông báo.');
      }

      const vapidResponse = await fetch('/api/push/vapid-key');
      const { publicKey } = await vapidResponse.json();

      const registration = await navigator.serviceWorker.ready;
      const subscription = await registration.pushManager.subscribe({
        userVisibleOnly: true,
        applicationServerKey: urlBase64ToUint8Array(publicKey)
      });

      // Save subscription to Firestore under user document
      await updateDoc(doc(db, 'users', userId), {
        pushSubscriptions: arrayUnion(JSON.parse(JSON.stringify(subscription)))
      });

      setIsSubscribed(true);
      return true;
    } catch (e) {
      console.error('Lỗi khi đăng ký push notification:', e);
      return false;
    }
  };

  const unsubscribeFromPush = async () => {
    if (!userId) return;
    try {
      const registration = await navigator.serviceWorker.ready;
      const subscription = await registration.pushManager.getSubscription();
      if (subscription) {
        await subscription.unsubscribe();
        await updateDoc(doc(db, 'users', userId), {
          pushSubscriptions: arrayRemove(JSON.parse(JSON.stringify(subscription)))
        });
      }
      setIsSubscribed(false);
    } catch (e) {
      console.error('Lỗi hủy Push notification:', e);
    }
  };

  return {
    isSupported,
    isSubscribed,
    permission,
    subscribeToPush,
    unsubscribeFromPush
  };
}
