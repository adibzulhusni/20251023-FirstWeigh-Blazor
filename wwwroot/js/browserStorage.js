window.browserStorage = {lastActivityTime: new Date().toISOString(),

    getItem: function (key) {
        return localStorage.getItem(key);
    },

    setItem: function (key, value) {
        localStorage.setItem(key, value);
    },

    removeItem: function (key) {
        localStorage.removeItem(key);
    },

    startActivityTracking: function () {
        const updateActivity = () => {
            browserStorage.lastActivityTime = new Date().toISOString();
        };

        // Track mouse movements, clicks, and keyboard events
        document.addEventListener('mousemove', updateActivity);
        document.addEventListener('mousedown', updateActivity);
        document.addEventListener('keydown', updateActivity);
        document.addEventListener('scroll', updateActivity);
        document.addEventListener('touchstart', updateActivity);
    },

    getLastActivityTime: function () {
        return browserStorage.lastActivityTime;
    }
};