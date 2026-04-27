(function (window, document) {
    'use strict';

    const overlayId = 'loadingOverlay';
    const frameId = 'bmsDownloadFrame';
    let fallbackTimer = null;

    function ensureOverlay() {
        let overlay = document.getElementById(overlayId);
        if (overlay) {
            return overlay;
        }

        overlay = document.createElement('div');
        overlay.id = overlayId;
        overlay.className = 'loading-overlay';
        overlay.setAttribute('data-bms-generated', 'true');
        overlay.innerHTML = `
            <div class="loading-content">
                <div class="loading-spinner"></div>
                <p class="loading-text">Loading...</p>
                <p class="loading-subtext">Please wait</p>
            </div>`;
        document.body.appendChild(overlay);
        return overlay;
    }

    function setText(overlay, message, subMessage) {
        const textEl = overlay.querySelector('.loading-text') || document.getElementById('loadingText');
        const subTextEl = overlay.querySelector('.loading-subtext') || document.getElementById('loadingSubtext');

        if (textEl) {
            textEl.textContent = message || 'Loading...';
        }

        if (subTextEl) {
            subTextEl.textContent = subMessage || 'Please wait';
        }
    }

    function show(message, subMessage) {
        const overlay = ensureOverlay();
        setText(overlay, message, subMessage);
        overlay.classList.add('active');
        document.body.style.overflow = 'hidden';
    }

    function hide() {
        const overlay = document.getElementById(overlayId);
        if (!overlay) {
            return;
        }

        overlay.classList.remove('active');
        document.body.style.overflow = '';

        if (overlay.getAttribute('data-bms-generated') === 'true') {
            window.setTimeout(function () {
                if (!overlay.classList.contains('active')) {
                    overlay.remove();
                }
            }, 300);
        }
    }

    function getFrame() {
        let frame = document.getElementById(frameId);
        if (!frame) {
            frame = document.createElement('iframe');
            frame.id = frameId;
            frame.name = frameId;
            frame.style.display = 'none';
            document.body.appendChild(frame);
        }
        return frame;
    }

    function inspectDownloadResponse(frame) {
        try {
            const doc = frame.contentDocument || frame.contentWindow.document;
            const bodyText = doc && doc.body ? doc.body.innerText.trim() : '';
            if (bodyText && bodyText.toLowerCase().indexOf('error') >= 0) {
                alert(bodyText.substring(0, 1000));
            }
        } catch (ignore) {
            // Attachment responses often cannot be inspected reliably; the download itself is enough.
        }
    }

    function download(url, message, subMessage) {
        const frame = getFrame();
        show(message || 'Preparing export...', subMessage || 'Please wait');

        if (fallbackTimer) {
            window.clearTimeout(fallbackTimer);
        }

        frame.onload = function () {
            inspectDownloadResponse(frame);
            window.setTimeout(hide, 400);
        };
        frame.onerror = function () {
            hide();
            alert('Unable to export data. Please try again.');
        };

        const separator = url.indexOf('?') >= 0 ? '&' : '?';
        fallbackTimer = window.setTimeout(hide, 120000);
        frame.src = url + separator + '_downloadTs=' + new Date().getTime();
    }

    window.BMSLoading = {
        show: show,
        hide: hide,
        download: download
    };
})(window, document);
