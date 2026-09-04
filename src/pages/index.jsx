import React, { useEffect } from 'react'
import { useRouter } from "next/router";

const index = () => {
  const router = useRouter();

  useEffect(() => {
    router.push('/count')
  }, []);

  return (
    <div>
      
    </div>
  )
}

export default index;
