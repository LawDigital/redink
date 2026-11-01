"""Request-independent HTTP/HTTPS listener lifecycle.

The HTTP worker never touches a listening socket. All lifecycle transitions
happen in the supervisor thread after the web response has been returned.
"""
from __future__ import annotations

import logging
import threading
import time
from typing import Callable

logger = logging.getLogger(__name__)


class ListenerSupervisor:
    def __init__(self, initial_settings, factory: Callable, on_applied: Callable,
                 on_failed: Callable, *, response_grace: float = .6):
        self._settings = initial_settings
        self._factory = factory
        self._on_applied = on_applied
        self._on_failed = on_failed
        self._response_grace = response_grace
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._wake = threading.Event()
        self._pending = None
        self._state = {'status': 'starting', 'message': 'Starting listener.'}
        self._listener = None
        self._worker = None
        self._worker_error = None
        self._intentional_stop = False

    def status(self) -> dict:
        with self._lock:
            return dict(self._state)

    def is_pending(self) -> bool:
        with self._lock:
            return self._pending is not None or self._state['status'] == 'applying'

    def reconfigure(self, settings, previous_configuration: str) -> bool:
        with self._lock:
            if self._stop_event.is_set() or self._pending is not None or self._state['status'] in ('applying', 'stopped'):
                return False
            self._pending = (settings, previous_configuration)
            self._state = {'status': 'pending', 'message': 'Applying listener and TLS settings without restarting the server process.'}
            self._wake.set()
        return True

    def shutdown(self) -> bool:
        if self._stop_event.is_set():
            return False
        self._stop_event.set()
        with self._lock:
            self._state = {'status': 'stopping', 'message': 'Gracefully stopping the listener.'}
        self._wake.set()
        return True

    def _serve(self, listener):
        try:
            listener.serve()
        except Exception as exc:
            if not self._intentional_stop:
                self._worker_error = exc
                self._wake.set()
            else:
                logger.debug('Transport worker exited during requested shutdown (%s).', type(exc).__name__)

    def _start(self, settings):
        listener = self._factory(settings)
        try:
            listener.prepare()  # Bind synchronously before reporting readiness.
        except Exception:
            try:
                listener.stop()
            except Exception:
                pass
            raise
        self._intentional_stop = False
        self._worker_error = None
        worker = threading.Thread(target=self._serve, args=(listener,), name='RedInkListener', daemon=False)
        self._listener = listener
        self._worker = worker
        worker.start()

    def _stop_listener(self):
        listener, worker = self._listener, self._worker
        if listener is None:
            return
        self._intentional_stop = True
        try:
            listener.stop()
        finally:
            if worker is not None and worker.is_alive():
                worker.join(timeout=15)
            if worker is not None and worker.is_alive():
                raise RuntimeError('Listener did not stop cleanly; refusing to start another on the same port.')
            self._listener = None
            self._worker = None

    def run(self, on_ready=None):
        try:
            self._start(self._settings)
            with self._lock:
                self._state = {'status': 'active', 'message': 'Listener is accepting connections.'}
            if on_ready:
                on_ready(self._settings)
            while not self._stop_event.is_set():
                self._wake.wait(.25)
                self._wake.clear()
                if self._stop_event.is_set():
                    break
                if self._worker_error is not None:
                    raise RuntimeError('Listener worker failed (' + type(self._worker_error).__name__ + ').')
                if self._worker is not None and not self._worker.is_alive():
                    raise RuntimeError('Listener stopped unexpectedly; refusing to report an active server.')
                with self._lock:
                    requested = self._pending
                    if requested is not None:
                        self._pending = None
                        self._state = {'status': 'applying', 'message': 'Switching the network listener.'}
                if requested is None:
                    continue
                updated, previous_configuration = requested
                # The response must leave the worker before its listening socket
                # is closed. The next listener is still run by this supervisor.
                if self._stop_event.wait(self._response_grace):
                    break
                previous_settings = self._settings
                try:
                    self._stop_listener()
                    self._start(updated)
                    self._settings = updated
                    self._on_applied(updated)
                    with self._lock:
                        self._state = {'status': 'active', 'message': 'Updated settings are active.'}
                    logger.info('Listener settings applied; protocol=%s; bind=%s:%d',
                                'https' if updated.tls_enabled else 'http', updated.host, updated.port)
                except Exception as exc:
                    logger.error('Listener update failed (%s). Restoring previous configuration.', type(exc).__name__)
                    error = type(exc).__name__
                    try:
                        if self._listener is not None:
                            self._stop_listener()
                        # This is intentionally performed even when binding fails;
                        # disk configuration must match the listener actually used.
                        self._on_failed(previous_settings, previous_configuration)
                        self._start(previous_settings)
                        self._settings = previous_settings
                        with self._lock:
                            self._state = {'status': 'failed', 'message': 'Listener update failed (' + error + '); previous settings restored.'}
                    except Exception as restoration_error:
                        logger.critical('Listener rollback failed (%s). Aborting the server.', type(restoration_error).__name__)
                        raise RuntimeError('Could not restore the previous listener; server stopped safely.') from restoration_error
        finally:
            self._stop_listener()
            with self._lock:
                self._state = {'status': 'stopped', 'message': 'Listener stopped.'}
