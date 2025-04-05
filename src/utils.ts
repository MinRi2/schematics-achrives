interface PollingOptions<T> {
    interval?: number;
    retryOnError?: Boolean;
    finished: (result: T) => Boolean,
}

function startPolling<T>(
    request: () => Promise<T>,
    options: PollingOptions<T> = {
        interval: 1000,
        retryOnError: true,
        finished: () => true,
    }
): Promise<T> {
    let timer: ReturnType<typeof setTimeout> | null = null;

    let resolve!: (value: T) => void;
    let reject!: (reason?: unknown) => void;

    const promise = new Promise<T>((res, rej) => {
        resolve = res;
        reject = rej;
    });

    const poll = async () => {
        try {
            const result = await request();
            if (options.finished(result)) {
                resolve(result);
                return;
            }

            timer = setTimeout(poll, options.interval);
        } catch (error) {
            if (options.retryOnError) {
                timer = setTimeout(poll, options.interval);
            } else {
                reject(error);
            }
        }
    };

    poll();

    return promise;
}

export { startPolling };