import redis

r = redis.Redis(host="127.0.0.1", port=6379)
r.set('K-0001', '111111')
print(r.get("K-0001"))