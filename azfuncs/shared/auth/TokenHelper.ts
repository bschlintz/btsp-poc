import * as jwt from 'jsonwebtoken';

export const getTokenFromAuthorizationHeader = (authorizationHeader: string): string => {
  const authHeaderParts = authorizationHeader.split(' ');
  return authHeaderParts.length === 2 ? authHeaderParts[1] : '';
}

export const getUpnFromToken = (token: string): string => {
  try {
    const decodedToken = jwt.decode(token);
    return decodedToken.upn;
  }
  catch (error) {
    throw error;
  }
}