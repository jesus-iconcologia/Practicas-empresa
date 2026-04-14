import json


def _json_progress(current, total, status='processing', message='', redirect_url=''):
    data = {
        'current': current,
        'total': total,
        'status': status,
        'message': message,
    }
    if redirect_url:
        data['redirect'] = redirect_url
    return json.dumps(data) + '\n'


def _should_update(i, total, min_updates=80):
    if total <= min_updates:
        return True
    step = max(1, total // min_updates)
    return (i + 1) % step == 0 or i + 1 == total